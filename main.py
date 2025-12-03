import os
import json
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from openai import OpenAI

# ==========================
# Carregar variáveis de ambiente
# ==========================
load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL não definido no .env")

if not OPENAI_API_KEY:
    raise RuntimeError("OPENAI_API_KEY não definido no .env")

# ==========================
# Conexão com banco e OpenAI
# ==========================
engine = create_engine(DATABASE_URL)
client = OpenAI(api_key=OPENAI_API_KEY)


# ==========================
# Login (email + senha)
# ==========================
def login():
    """
    Tela de login via terminal:
    - Email
    - Senha = valor da coluna 'id' da tabela person
    Retorna: (email, nome, area, cargo)
    """
    email = input("Digite seu email: ").strip().lower()
    senha = input("Digite sua senha: ").strip()

    with engine.connect() as conn:
        query = text(
            """
            SELECT id, nome, area, posicao
            FROM person
            WHERE email = :email
            LIMIT 1
            """
        )
        df = pd.read_sql(query, conn, params={"email": email})

    if df.empty:
        print("❌ Email não encontrado.")
        raise SystemExit(1)

    row = df.iloc[0]
    id_banco = str(row["id"])

    if senha != id_banco:
        print("❌ Senha incorreta.")
        raise SystemExit(1)

    nome = row["nome"]
    area = row["area"]
    cargo = row["posicao"]

    print("\n✅ Login realizado com sucesso!")
    print(f"Usuário: {nome} | Área: {area} | Cargo: {cargo}\n")

    return email, nome, area, cargo


# ==========================
# Escolha de período
# ==========================
def escolher_periodo():
    """
    Pergunta se vai usar calendário do Teams e qual período.
    Retorna:
        usar_teams (bool),
        periodo_data_inicial,
        periodo_data_final
    """
    periodo_data_inicial = None
    periodo_data_final = None

    usar_teams = input("Você quer usar o calendário do Teams? (Sim/Não): ").strip().lower()

    if usar_teams == "sim":
        print("\nEscolha o período:")
        print("1 - Últimos 7 dias")
        print("2 - Esta semana (segunda a sexta)")
        print("3 - Semana anterior (segunda a sexta)")
        opcao = input("Digite 1, 2 ou 3: ").strip()

        hoje = dt.date.today()

        if opcao == "1":
            periodo_data_final = hoje
            periodo_data_inicial = hoje - dt.timedelta(days=7)

        elif opcao == "2":
            weekday = hoje.weekday()  # Monday = 0
            segunda = hoje - dt.timedelta(days=weekday)
            sexta = segunda + dt.timedelta(days=4)
            periodo_data_inicial = segunda
            periodo_data_final = sexta

        elif opcao == "3":
            weekday = hoje.weekday()
            segunda_atual = hoje - dt.timedelta(days=weekday)
            segunda_passada = segunda_atual - dt.timedelta(days=7)
            sexta_passada = segunda_passada + dt.timedelta(days=4)
            periodo_data_inicial = segunda_passada
            periodo_data_final = sexta_passada

        else:
            print("Opção inválida.")
            raise SystemExit(1)

        print("\n✅ Período selecionado:")
        print("Data inicial:", periodo_data_inicial)
        print("Data final  :", periodo_data_final)
        return True, periodo_data_inicial, periodo_data_final

    else:
        print("Você optou por não usar o calendário do Teams.")
        return False, periodo_data_inicial, periodo_data_final


# ==========================
# Query combinados
# ==========================
def buscar_combinados(email, periodo_data_inicial):
    """
    Busca combinados da tabela 'combinados' com as regras:
    - employee_key = email
    - E (status = 'started'
       OU (status = 'completed' E status_assigned_at > periodo_data_inicial))
    Retorna:
        df_combinados, combinados_texto
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

    print("Registros encontrados em 'combinados':", len(combinados))

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


# ==========================
# Microsoft Graph – Token e reuniões
# ==========================
def get_access_token():
    """
    Gera um access token do Azure AD para chamar o Microsoft Graph
    usando client_credentials.
    """
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data)
    if resp.status_code != 200:
        raise Exception(f"Erro ao obter token: {resp.status_code} - {resp.text}")
    return resp.json()["access_token"]


def get_calendar_events(email, data_inicial, data_final):
    """
    Busca eventos de calendário de um usuário (email/UPN) no intervalo indicado.
    email: e-mail do usuário no tenant, ex: 'pessoa@mindsight.com.br'
    data_inicial, data_final: objetos datetime.date
    """
    access_token = get_access_token()

    # Monta datas em ISO 8601 (início 00:00, fim 23:59)
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
            raise Exception(f"Erro ao buscar eventos: {resp.status_code} - {resp.text}")

        data = resp.json()
        events.extend(data.get("value", []))

        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        # Paginação: próxima página
        url = next_link
        params = None  # já vem tudo na URL

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

    df = pd.DataFrame(linhas)
    return df


# ==========================
# Ajuste de fuso e texto das reuniões
# ==========================
def ajustar_horarios_brasilia(df):
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


def gerar_texto_reunioes(df):
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

    texto_reunioes = " ".join(linhas)
    return texto_reunioes


# ==========================
# OpenAI – resumo com reuniões + combinados
# ==========================
def gerar_resumo_com_base_em_reunioes_comb(
    texto_reunioes, texto_reunioes_anterior, combinados_texto, nome, cargo, area, model="gpt-4o-mini"
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
        f"Faça o resumo de trabalho de um funcionário {nome} tentando dizer no que o funcionário trabalhou "
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
# main()
# ==========================
def main():
    # 1) Login
    email, nome, area, cargo = login()

    # 2) Escolher período (Teams)
    usar_teams, periodo_data_inicial, periodo_data_final = escolher_periodo()

    if not usar_teams:
        print("Por enquanto o script depende do calendário do Teams para gerar o resumo.")
        raise SystemExit(0)

    # 3) Período anterior (comparação)
    periodo_data_inicial_anterior = periodo_data_inicial - dt.timedelta(days=7)
    periodo_data_final_anterior = periodo_data_final - dt.timedelta(days=7)

    print("\nPeríodo anterior:")
    print("Data inicial anterior:", periodo_data_inicial_anterior)
    print("Data final anterior  :", periodo_data_final_anterior)

    # 4) Buscar combinados
    _, combinados_texto = buscar_combinados(email, periodo_data_inicial)

    # 5) Buscar reuniões (período atual e anterior)
    user_id = email
    usuario_display_name = nome
    usuario_upn = email

    reunioes_atual = get_calendar_events(email, periodo_data_inicial, periodo_data_final)
    print(f"\nTotal de reuniões encontradas (período atual): {len(reunioes_atual)}")

    df_reunioes = eventos_para_dataframe_v2(
        reunioes_atual,
        user_id=user_id,
        usuario_display_name=usuario_display_name,
        usuario_upn=usuario_upn,
    )

    reunioes_anterior = get_calendar_events(
        email, periodo_data_inicial_anterior, periodo_data_final_anterior
    )
    print(f"Total de reuniões encontradas (período anterior): {len(reunioes_anterior)}")

    df_reunioes_anterior = eventos_para_dataframe_v2(
        reunioes_anterior,
        user_id=user_id,
        usuario_display_name=usuario_display_name,
        usuario_upn=usuario_upn,
    )

    # 6) Ajustar fuso e gerar textos
    df_reunioes_br = ajustar_horarios_brasilia(df_reunioes)
    df_reunioes_anterior_br = ajustar_horarios_brasilia(df_reunioes_anterior)

    texto_reunioes = gerar_texto_reunioes(df_reunioes_br)
    texto_reunioes_anterior = gerar_texto_reunioes(df_reunioes_anterior_br)

    # 7) Chamar OpenAI para gerar resumo
    respostas_json, raw_text = gerar_resumo_com_base_em_reunioes_comb(
        texto_reunioes,
        texto_reunioes_anterior,
        combinados_texto,
        nome,
        cargo,
        area,
    )

    print("\n==== RESPOSTA CRUA DO MODELO ====\n")
    print(raw_text)

    if respostas_json is not None:
        print("\n==== JSON INTERPRETADO ====\n")
        print(json.dumps(respostas_json, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
