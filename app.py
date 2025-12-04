import os
import json
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from openai import OpenAI

# ==========================
# Configuração inicial
# ==========================
st.set_page_config(page_title="Resumo Calendário Teams", page_icon="📅")

load_dotenv()  # localmente; no Railway as vars vêm do painel

DATABASE_URL = os.getenv("DATABASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

if not DATABASE_URL:
    st.error("DATABASE_URL não definido nas variáveis de ambiente.")
    st.stop()

if not OPENAI_API_KEY:
    st.error("OPENAI_API_KEY não definido nas variáveis de ambiente.")
    st.stop()

if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
    st.error("TENANT_ID, CLIENT_ID ou CLIENT_SECRET não definidos.")
    st.stop()

engine = create_engine(DATABASE_URL)
client = OpenAI(api_key=OPENAI_API_KEY)

# ==========================
# Funções auxiliares
# ==========================

def login_db(email: str, senha: str):
    """
    Valida email + senha (id da tabela person).
    Retorna (email_norm, nome, area, cargo).
    """
    email_norm = email.strip().lower()
    senha = senha.strip()

    with engine.connect() as conn:
        query = text(
            """
            SELECT id, nome, area, posicao
            FROM person
            WHERE email = :email
            LIMIT 1
            """
        )
        df = pd.read_sql(query, conn, params={"email": email_norm})

    if df.empty:
        raise ValueError("Email não encontrado na tabela person.")

    row = df.iloc[0]
    id_banco = str(row["id"])

    if senha != id_banco:
        raise ValueError("Senha incorreta.")

    nome = row["nome"]
    area = row["area"]
    cargo = row["posicao"]

    return email_norm, nome, area, cargo


def escolher_periodo(periodo_opcao: int):
    """
    1 = últimos 7 dias
    2 = semana atual (seg-sex)
    3 = semana anterior (seg-sex)
    """
    hoje = dt.date.today()

    if periodo_opcao == 1:
        periodo_data_final = hoje
        periodo_data_inicial = hoje - dt.timedelta(days=7)
    elif periodo_opcao == 2:
        weekday = hoje.weekday()  # Monday = 0
        segunda = hoje - dt.timedelta(days=weekday)
        sexta = segunda + dt.timedelta(days=4)
        periodo_data_inicial = segunda
        periodo_data_final = sexta
    elif periodo_opcao == 3:
        weekday = hoje.weekday()
        segunda_atual = hoje - dt.timedelta(days=weekday)
        segunda_passada = segunda_atual - dt.timedelta(days=7)
        sexta_passada = segunda_passada + dt.timedelta(days=4)
        periodo_data_inicial = segunda_passada
        periodo_data_final = sexta_passada
    else:
        raise ValueError("Opção de período inválida. Use 1, 2 ou 3.")

    return periodo_data_inicial, periodo_data_final


def buscar_combinados(email: str, periodo_data_inicial: dt.date):
    """
    Busca combinados da tabela 'combinados' com as regras:
    - employee_key = email
    - E (status = 'started'
       OU (status = 'completed' E status_assigned_at > periodo_data_inicial))
    """
    if periodo_data_inicial is None:
        raise ValueError("periodo_data_inicial não pode ser None para filtrar combinados.")

    with engine.connect() as conn:
        query_combinados = text(
            """
            SELECT *
            FROM combinados
            WHERE employee_key = :email
              AND (
                    status = 'started'
                    OR (
                        status = 'completed'
                        AND status_assigned_at > :periodo_data_inicial
                    )
                  )
            """
        )

        combinados = pd.read_sql(
            query_combinados,
            conn,
            params={
                "email": email,
                "periodo_data_inicial": periodo_data_inicial,
            },
        )

    if combinados.empty:
        combinados_texto = "Nenhum combinado encontrado para este período."
    else:
        textos = []
        for _, row in combinados.iterrows():
            nome_combinado = row.get("name", "")
            descricao = row.get("description", "")
            due_date = row.get("due_date", "")
            status = row.get("status", "")
            prioridade = row.get("priority", "")
            modified = row.get("modified", "")

            texto = (
                f"O combinado é {nome_combinado}, que tem como descrição {descricao}. "
                f"Que deve ser feito até {due_date} e tem o status {status}, "
                f"com prioridade {prioridade}, e sua última atualização foi {modified}."
            )
            textos.append(texto)

        combinados_texto = " ".join(textos)

    return combinados, combinados_texto


def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data)
    if resp.status_code != 200:
        raise RuntimeError(f"Erro ao obter token do Graph: {resp.status_code} - {resp.text}")
    return resp.json()["access_token"]


def get_calendar_events(email: str, data_inicial: dt.date, data_final: dt.date):
    access_token = get_access_token()

    start_dt = dt.datetime.combine(data_inicial, dt.time(hour=0, minute=0))
    end_dt = dt.datetime.combine(data_final, dt.time(hour=23, minute=59))

    start_iso = start_dt.isoformat()
    end_iso = end_dt.isoformat()

    url = f"https://graph.microsoft.com/v1.0/users/{email}/calendarView"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    params = {"startDateTime": start_iso, "endDateTime": end_iso, "$orderby": "start/dateTime"}

    events = []
    while True:
        resp = requests.get(url, headers=headers, params=params)
        if resp.status_code != 200:
            raise RuntimeError(f"Erro ao buscar eventos do calendário: {resp.status_code} - {resp.text}")

        data = resp.json()
        events.extend(data.get("value", []))

        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        url = next_link
        params = None

    return events


def eventos_para_dataframe_v2(events, user_id=None, usuario_display_name=None, usuario_upn=None):
    linhas = []

    for ev in events:
        subject = ev.get("subject", "(sem assunto)")
        organizer = ev.get("organizer", {}).get("emailAddress", {})
        organizador_nome = organizer.get("name", "")
        organizador_email = organizer.get("address", "")

        start = ev.get("start", {})
        end = ev.get("end", {})

        inicio_raw = start.get("dateTime")
        fim_raw = end.get("dateTime")
        inicio_tz = start.get("timeZone")
        fim_tz = end.get("timeZone")

        inicio_dt = pd.to_datetime(inicio_raw) if inicio_raw else None
        fim_dt = pd.to_datetime(fim_raw) if fim_raw else None

        attendees = ev.get("attendees", [])
        participantes_emails = []
        participantes_nomes = []

        for a in attendees:
            email_info = a.get("emailAddress", {})
            addr = email_info.get("address", "")
            name = email_info.get("name", "")
            if addr:
                participantes_emails.append(addr)
            if name:
                participantes_nomes.append(name)

        participantes_emails_str = "; ".join(participantes_emails) if participantes_emails else ""
        participantes_nomes_str = "; ".join(participantes_nomes) if participantes_nomes else ""

        local = ev.get("location", {}).get("displayName", "")

        reuniao_online = False
        link_reuniao_online = ""

        if ev.get("isOnlineMeeting") is True:
            reuniao_online = True
            link_reuniao_online = (
                ev.get("onlineMeetingUrl", "") or ev.get("joinUrl", "") or ""
            )

        if not link_reuniao_online:
            online_meeting = ev.get("onlineMeeting", {})
            if isinstance(online_meeting, dict):
                link_reuniao_online = online_meeting.get("joinUrl", "") or link_reuniao_online
                if link_reuniao_online:
                    reuniao_online = True

        linhas.append(
            {
                "user_id": user_id,
                "assunto": subject,
                "organizador_nome": organizador_nome,
                "organizador_email": organizador_email,
                "inicio_datetime": inicio_dt,
                "inicio_timezone": inicio_tz,
                "fim_datetime": fim_dt,
                "fim_timezone": fim_tz,
                "participantes_emails": participantes_emails_str,
                "participantes_nomes": participantes_nomes_str,
                "local": local,
                "reuniao_online": reuniao_online,
                "link_reuniao_online": link_reuniao_online,
                "usuario_display_name": usuario_display_name,
                "usuario_upn": usuario_upn,
            }
        )

    return pd.DataFrame(linhas)


def ajustar_horarios_brasilia(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    tz_br = ZoneInfo("America/Sao_Paulo")

    def converter_para_brasilia(x):
        if pd.isna(x):
            return x
        if x.tzinfo is None:
            x = x.replace(tzinfo=dt.timezone.utc)
        return x.astimezone(tz_br)

    df["inicio_datetime"] = df["inicio_datetime"].apply(converter_para_brasilia)
    df["fim_datetime"] = df["fim_datetime"].apply(converter_para_brasilia)

    if "inicio_timezone" in df.columns:
        df["inicio_timezone"] = "America/Sao_Paulo"
    if "fim_timezone" in df.columns:
        df["fim_timezone"] = "America/Sao_Paulo"

    return df


def gerar_texto_reunioes(df: pd.DataFrame) -> str:
    linhas = []

    for _, row in df.iterrows():
        assunto = row.get("assunto", "(sem assunto)") or "(sem assunto)"
        inicio = row.get("inicio_datetime", None)
        fim = row.get("fim_datetime", None)

        if pd.notna(inicio):
            inicio_str = inicio.strftime("%d/%m/%Y %H:%M")
        else:
            inicio_str = "sem horário de início"

        if pd.notna(fim):
            fim_str = fim.strftime("%d/%m/%Y %H:%M")
        else:
            fim_str = "sem horário de término"

        linha = (
            f"{assunto}: horário de início: {inicio_str} "
            f"e término {fim_str}. Teve a reunião: {assunto}."
        )
        linhas.append(linha)

    return " ".join(linhas)


def gerar_resumo_com_base_em_reunioes_comb(
    texto_reunioes: str,
    texto_reunioes_anterior: str,
    combinados_texto: str,
    nome: str,
    cargo: str,
    area: str,
    model: str = "gpt-4o-mini",
):
    system_prompt = (
        "Você é um assistente que ajuda funcionários a preencher o resumo de tarefas "
        "com base nas reuniões do calendário.\n\n"
        "REGRAS:\n"
        "1. Use APENAS as informações fornecidas nas reuniões (texto_reunioes).\n"
        "2. Não invente reuniões e não faça grandes inferências.\n"
        "3. Faça como um texto corrido, evitando apenas escrever o calendário.\n"
        "4. Comece sempre com (Resumo da semana:) e em seguida o resumo gerado.\n"
        "5. Tente resumir com as principais reuniões e trazendo apenas informações profissionais, "
        "não colocando no resumo eventos pessoais ou médicos ou de rotina, como almoço e atividades físicas, "
        "nem cite coisas que não são importantes para a profissão."
    )

    user_prompt = (
        f"Faça o resumo de trabalho de um funcionário {nome} tentando dizer no que o funcionário trabalhou em primeira pessoa, ou seja, como se a pessoa estivesse escrevendo. "
        f"e quais atividades a pessoa se dedicou mais. "
        f"Leve em consideração o cargo da pessoa: {cargo} na área {area}. "
        f"O calendário dessa semana da pessoa foi: {texto_reunioes}. "
        f"Compare com a semana passada para tentar ver projetos que estão com mais foco nesta semana "
        f"e na semana passada e projetos que podem ter perdido prioridade: {texto_reunioes_anterior}. "
        f"Leve em consideração os pedidos do líder da pessoa com o funcionário, que são os seguintes: {combinados_texto}. "
        "Lembre-se: não invente nada. Caso não tenha informação, retorne apenas "
        "\"Sem informações suficientes.\".\n"
        "Retorne SOMENTE o JSON, sem nenhum texto extra."
    )

    completion = client.chat.completions.create(
        model=model,
        temperature=0.1,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )

    raw_text = completion.choices[0].message.content.strip()

    try:
        respostas_json = json.loads(raw_text)
    except json.JSONDecodeError:
        respostas_json = None

    return respostas_json, raw_text


# ==========================
# UI com Streamlit
# ==========================

def pagina_login():
    st.title("📅 Resumo do Calendário (Teams)")

    st.subheader("Login")
    with st.form("login_form"):
        email = st.text_input("Email corporativo")
        senha = st.text_input("Senha (ID da tabela person)", type="password")
        submit = st.form_submit_button("Entrar")

    if submit:
        try:
            email_norm, nome, area, cargo = login_db(email, senha)
            st.session_state.logged_in = True
            st.session_state.email = email_norm
            st.session_state.nome = nome
            st.session_state.area = area
            st.session_state.cargo = cargo
            st.success(f"Bem-vindo, {nome}!")
            st.rerun()
        except Exception as e:
            st.error(f"Erro no login: {e}")


def pagina_principal():
    st.title("📅 Resumo do Calendário (Teams)")

    nome = st.session_state.nome
    area = st.session_state.area
    cargo = st.session_state.cargo
    email = st.session_state.email

    st.markdown(f"**Usuário:** {nome}  \n**Área:** {area}  \n**Cargo:** {cargo}  \n**Email:** {email}")

    if st.button("Sair"):
        for k in ["logged_in", "email", "nome", "area", "cargo"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    st.markdown("---")
    st.subheader("Configurações do resumo")

    opcao_texto = {
        "Últimos 7 dias": 1,
        "Esta semana (seg-sex)": 2,
        "Semana anterior (seg-sex)": 3,
    }

    escolha = st.selectbox(
        "Período para considerar",
        list(opcao_texto.keys()),
        index=1,  # default: esta semana
    )
    periodo_opcao = opcao_texto[escolha]

    if st.button("Gerar resumo"):
        try:
            with st.spinner("Gerando resumo com base no calendário e combinados..."):
                # Período atual e anterior
                periodo_data_inicial, periodo_data_final = escolher_periodo(periodo_opcao)
                periodo_data_inicial_anterior = periodo_data_inicial - dt.timedelta(days=7)
                periodo_data_final_anterior = periodo_data_final - dt.timedelta(days=7)

                # Combinados
                _, combinados_texto = buscar_combinados(email, periodo_data_inicial)

                # Reuniões
                reunioes_atual = get_calendar_events(email, periodo_data_inicial, periodo_data_final)
                df_reunioes = eventos_para_dataframe_v2(
                    reunioes_atual,
                    user_id=email,
                    usuario_display_name=nome,
                    usuario_upn=email,
                )

                reunioes_anterior = get_calendar_events(
                    email, periodo_data_inicial_anterior, periodo_data_final_anterior
                )
                df_reunioes_anterior = eventos_para_dataframe_v2(
                    reunioes_anterior,
                    user_id=email,
                    usuario_display_name=nome,
                    usuario_upn=email,
                )

                # Ajustar fuso + gerar textos
                df_reunioes_br = ajustar_horarios_brasilia(df_reunioes)
                df_reunioes_anterior_br = ajustar_horarios_brasilia(df_reunioes_anterior)

                texto_reunioes = gerar_texto_reunioes(df_reunioes_br)
                texto_reunioes_anterior = gerar_texto_reunioes(df_reunioes_anterior_br)

                # Chamar OpenAI
                respostas_json, raw_text = gerar_resumo_com_base_em_reunioes_comb(
                    texto_reunioes,
                    texto_reunioes_anterior,
                    combinados_texto,
                    nome,
                    cargo,
                    area,
                )

            st.markdown(
                f"""
                <div style="
                    background-color:#f8f9fa;
                    border-left: 5px solid #4a90e2;
                    padding: 1rem;
                    border-radius: 8px;
                    font-size: 1rem;
                    line-height: 1.6;
                    color:#333;
                    white-space: pre-wrap;
                ">
                    {raw_text}
                </div>
                """,
                unsafe_allow_html=True
            )

        except Exception as e:
            st.error(f"Erro ao gerar resumo: {e}")


def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        pagina_login()
    else:
        pagina_principal()


if __name__ == "__main__":
    main()
