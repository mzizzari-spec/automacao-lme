"""
Script: carga_inicial.py
Função: Roda UMA VEZ SÓ para carregar os últimos 3 meses de dados históricos
        do site shockmetais.com.br/lme para o Google Sheets.
ATENÇÃO: Após rodar, não precisa rodar de novo. Use o coletar_lme.py diariamente.
"""

import os
import json
import calendar
from datetime import datetime, date, timedelta

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials


# ─────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
URL_SITE = "https://shockmetais.com.br/lme"

CABECALHO = [
    "Data", "Dia da Semana", "Tipo",
    "Cobre (US$/t)", "Alumínio (US$/t)", "Dólar (R$/US$)",
    "Cobre (R$/kg)", "Alumínio (R$/kg)", "Mês"
]

CABECALHO_SEM_MES = CABECALHO[:-1]

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

FERIADOS_FIXOS = {
    (1, 1), (21, 4), (1, 5), (7, 9),
    (12, 10), (2, 11), (15, 11), (25, 12),
}

def calcular_pascoa(ano):
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(ano, mes, dia)

def feriados_moveis(ano):
    pascoa = calcular_pascoa(ano)
    return {
        pascoa - timedelta(days=48),
        pascoa - timedelta(days=47),
        pascoa - timedelta(days=2),
        pascoa + timedelta(days=60),
    }

def eh_dia_util(d):
    if d.weekday() >= 5:
        return False
    if (d.day, d.month) in FERIADOS_FIXOS:
        return False
    if date(d.year, d.month, d.day) in feriados_moveis(d.year):
        return False
    return True

def dias_uteis_do_mes(ano, mes):
    _, ultimo_dia = calendar.monthrange(ano, mes)
    return [date(ano, mes, d) for d in range(1, ultimo_dia + 1)
            if eh_dia_util(date(ano, mes, d))]

def nome_aba(ano, mes):
    return f"{MESES_PT[mes-1]}/{ano}"

def conectar_google_sheets():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("Secret GOOGLE_CREDENTIALS_JSON não encontrado!")
    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)
        
def limpar_numero(texto, inteiro=False):
    if not texto or texto.strip().lower() in ("feriado", "-", ""):
        return None
    print(f"DEBUG limpar_numero: '{texto}'")
    try:
        s = texto.strip().replace(",", ".")
        valor = float(s)
        return int(round(valor * 1000)) if inteiro else valor
    except ValueError:
        return None

def para_float(valor):
    if not valor:
        return None
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return None

def calc_kg(usd_t, dolar):
    if usd_t and dolar:
        return round((usd_t * dolar) / 1000, 4)
    return None

def buscar_dados_mes(ano, mes):
    """Busca todos os dados de um mês no site."""
    mes_param = f"{MESES_PT[mes-1]}/{ano}"
    print(f"\n📅 Buscando dados de {mes_param}...")

    # O site usa parâmetro de mês na URL
    mes_url = f"{MESES_PT[mes-1].lower()}{ano}"
    url = f"{URL_SITE}?mes={mes-1}&ano={ano}"

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

    # Tenta buscar a página do mês
    response = requests.get(URL_SITE, headers=headers, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    # Verifica se há seletor de mês e tenta navegar
    # O site carrega o mês atual por padrão, para outros meses usa JS
    # Vamos buscar via URL com parâmetros
    meses_en = ["jan", "fev", "mar", "abr", "mai", "jun",
                "jul", "ago", "set", "out", "nov", "dez"]

    # Tenta URL alternativa
    url_mes = f"{URL_SITE}/{meses_en[mes-1]}-{ano}"
    try:
        response2 = requests.get(url_mes, headers=headers, timeout=30)
        if response2.status_code == 200:
            soup = BeautifulSoup(response2.text, "html.parser")
    except:
        pass

    tabela = soup.find("table")
    if not tabela:
        print(f"  ⚠️  Tabela não encontrada para {mes_param}")
        return {}

    dados_mes = {}
    dias_semana_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

    for linha in tabela.find_all("tr"):
        colunas = linha.find_all("td")
        if not colunas or len(colunas) < 8:
            continue

        dia_texto = colunas[0].get_text(strip=True)

        # Ignora linhas de média
        if "média" in dia_texto.lower() or "media" in dia_texto.lower():
            continue

        # Extrai o número do dia (ex: "01/Jun" → 1)
        import re
        match = re.match(r"(\d+)/", dia_texto)
        if not match:
            continue

        dia_num = int(match.group(1))
        try:
            d = date(ano, mes, dia_num)
        except ValueError:
            continue

        cobre = limpar_numero(colunas[1].get_text(strip=True), inteiro=True)
        aluminio = limpar_numero(colunas[3].get_text(strip=True), inteiro=True)
        dolar = limpar_numero(colunas[7].get_text(strip=True))

        if cobre or aluminio or dolar:
            dados_mes[d] = {
                "data": d.strftime("%d/%m/%Y"),
                "dia_semana": dias_semana_nomes[d.weekday()],
                "cobre": cobre,
                "aluminio": aluminio,
                "dolar": dolar,
            }
            print(f"  ✅ {d.strftime('%d/%m/%Y')} — Cobre: {cobre} | Alumínio: {aluminio} | Dólar: {dolar}")

    return dados_mes

def processar_mes(planilha, ano, mes, dados_reais):
    """Cria/atualiza aba do mês com dados reais + projeções + médias."""
    nome = nome_aba(ano, mes)
    hoje = datetime.now().date()

    # Cria ou limpa aba
    try:
        aba = planilha.worksheet(nome)
        aba.clear()
        print(f"  Aba '{nome}' limpa e será reescrita.")
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome, rows=200, cols=10)
        print(f"  Aba '{nome}' criada.")

    dias_uteis = dias_uteis_do_mes(ano, mes)

    # Acha último dia real disponível
    ultimo_real = None
    for d in sorted(dados_reais.keys(), reverse=True):
        ultimo_real = dados_reais[d]
        break

    linhas = []
    for d in dias_uteis:
        if d in dados_reais:
            r = dados_reais[d]
            linhas.append([
                r["data"], r["dia_semana"], "Real",
                r["cobre"], r["aluminio"], r["dolar"],
                calc_kg(r["cobre"], r["dolar"]),
                calc_kg(r["aluminio"], r["dolar"]),
            ])
        elif ultimo_real and d > hoje:
            linhas.append([
                d.strftime("%d/%m/%Y"),
                ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"][d.weekday()],
                "Projetado",
                ultimo_real["cobre"], ultimo_real["aluminio"], ultimo_real["dolar"],
                calc_kg(ultimo_real["cobre"], ultimo_real["dolar"]),
                calc_kg(ultimo_real["aluminio"], ultimo_real["dolar"]),
            ])

    # Médias
    def media(linhas_filtro, col):
        vals = [para_float(l[col]) for l in linhas_filtro if l[col] not in (None, "")]
        vals = [v for v in vals if v is not None]
        return round(sum(vals) / len(vals), 4) if vals else None

    reais = [l for l in linhas if l[2] == "Real"]
    todos = linhas

    media_real = ["Média Real", "", "", media(reais, 3), media(reais, 4), media(reais, 5), media(reais, 6), media(reais, 7)]
    media_proj = ["Média Projetada", "", "", media(todos, 3), media(todos, 4), media(todos, 5), media(todos, 6), media(todos, 7)]

    todas_linhas = [CABECALHO_SEM_MES] + linhas + [[], media_real, media_proj]
    aba.update("A1", todas_linhas)

    print(f"  ✅ {len(reais)} dias reais + {len(todos)-len(reais)} projetados gravados em '{nome}'")
    return linhas, nome

def atualizar_consolidado(planilha, todos_os_dados):
    """Cria/atualiza aba Consolidado com todos os meses."""
    try:
        aba = planilha.worksheet("Consolidado")
        aba.clear()
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title="Consolidado", rows=2000, cols=10)

    linhas_consolidado = [CABECALHO]
    for nome_mes, linhas in todos_os_dados:
        for l in linhas:
            linhas_consolidado.append(l + [nome_mes])

    aba.update("A1", linhas_consolidado)
    print(f"\n✅ Aba Consolidado atualizada com {len(linhas_consolidado)-1} linhas.")

def main():
    print("=" * 55)
    print("CARGA INICIAL — Últimos 3 meses")
    print("=" * 55)

    hoje = datetime.now()
    client = conectar_google_sheets()
    planilha = client.open_by_key(GOOGLE_SHEET_ID)

    todos_os_dados = []

    for i in range(2, -1, -1):  # 3 meses atrás até o atual
        # Calcula mês alvo
        mes_alvo = hoje.month - i
        ano_alvo = hoje.year
        while mes_alvo <= 0:
            mes_alvo += 12
            ano_alvo -= 1

        dados_reais = buscar_dados_mes(ano_alvo, mes_alvo)

        if not dados_reais and i > 0:
            print(f"  ⚠️  Nenhum dado encontrado, pulando.")
            continue

        linhas, nome = processar_mes(planilha, ano_alvo, mes_alvo, dados_reais)
        todos_os_dados.append((nome, linhas))

    atualizar_consolidado(planilha, todos_os_dados)

    print("\n" + "=" * 55)
    print("✅ Carga inicial concluída!")
    print("=" * 55)

if __name__ == "__main__":
    main()
