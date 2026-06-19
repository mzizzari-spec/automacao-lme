"""
Script: coletar_lme.py
Função: Coleta dados de Cobre, Alumínio e Dólar do site shockmetais.com.br/lme
        e grava na planilha do Google Sheets automaticamente.
Roda: Todo dia às 7h via GitHub Actions
"""

import os
import re
import json
from datetime import datetime, timedelta, date
import calendar

import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials


GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
URL_SITE = "https://shockmetais.com.br/lme"

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

CABECALHO = [
    "Data", "Dia da Semana", "Tipo",
    "Cobre (US$/t)", "Alumínio (US$/t)", "Dólar (R$/US$)",
    "Cobre (R$/kg)", "Alumínio (R$/kg)",
]

CABECALHO_CONSOLIDADO = CABECALHO + ["Mês"]

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

def limpar_numero(texto, americano=False):
    if not texto or texto.strip().lower() in ("feriado", "-", ""):
        return None
    try:
        s = texto.strip()
        if americano:
            # Formato americano: 13,690.00 → remove vírgula de milhar
            s = s.replace(",", "")
        else:
            # Formato brasileiro: 5,0923 → troca vírgula por ponto
            s = s.replace(".", "").replace(",", ".")
        return float(s)
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

def obter_dados_ontem():
    hoje = datetime.now()
    ontem = hoje - timedelta(days=1)
    if hoje.weekday() == 0:
        ontem = hoje - timedelta(days=3)

    mes_en_pt = {
        "Jan": "Jan", "Feb": "Fev", "Mar": "Mar", "Apr": "Abr",
        "May": "Mai", "Jun": "Jun", "Jul": "Jul", "Aug": "Ago",
        "Sep": "Set", "Oct": "Out", "Nov": "Nov", "Dec": "Dez"
    }
    mes_abrev = ontem.strftime("%b")
    dia_site = f"{ontem.day:02d}/{mes_en_pt.get(mes_abrev, mes_abrev)}"

    print(f"Buscando dados para: {dia_site} ({ontem.strftime('%d/%m/%Y')})")

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    response = requests.get(URL_SITE, headers=headers, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    tabela = soup.find("table")
    if not tabela:
        raise ValueError("Tabela não encontrada no site!")

    dias_semana_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

    for linha in tabela.find_all("tr"):
        colunas = linha.find_all("td")
        if not colunas:
            continue
        dia_celula = colunas[0].get_text(strip=True)
        if dia_site in dia_celula:
        cobre = limpar_numero(colunas[1].get_text(strip=True), americano=True)
        aluminio = limpar_numero(colunas[3].get_text(strip=True), americano=True)
        dolar = limpar_numero(colunas[7].get_text(strip=True))
            return {
                "data": ontem.strftime("%d/%m/%Y"),
                "data_obj": ontem.date(),
                "dia_semana": dias_semana_nomes[ontem.weekday()],
                "cobre_usd_t": cobre,
                "aluminio_usd_t": aluminio,
                "dolar_brl": dolar,
            }

    print(f"⚠️  Dia '{dia_site}' não encontrado (feriado ou fim de semana).")
    return None

def obter_ou_criar_aba(planilha, ano, mes):
    nome = nome_aba(ano, mes)
    try:
        aba = planilha.worksheet(nome)
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome, rows=200, cols=10)
        aba.append_row(CABECALHO)
        print(f"Aba '{nome}' criada.")
    return aba

def recalcular_aba(aba, ano, mes):
    """Recalcula projeções e médias da aba do mês."""
    todos_registros = aba.get_all_values()
    if not todos_registros:
        return []

    linhas_reais = [r for r in todos_registros[1:] if r and len(r) > 2 and r[2] == "Real"]
    if not linhas_reais:
        return []

    ultimo = linhas_reais[-1]
    ultimo_cobre = para_float(ultimo[3])
    ultimo_aluminio = para_float(ultimo[4])
    ultimo_dolar = para_float(ultimo[5])

    datas_reais = set(r[0] for r in linhas_reais)
    dias_uteis = dias_uteis_do_mes(ano, mes)
    dias_semana_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
    hoje = datetime.now().date()

    todas_linhas = []
    for d in dias_uteis:
        data_str = d.strftime("%d/%m/%Y")
        if data_str in datas_reais:
            linha_real = next(r for r in linhas_reais if r[0] == data_str)
            todas_linhas.append(linha_real[:8])
        elif d > hoje:
            todas_linhas.append([
                data_str,
                dias_semana_nomes[d.weekday()],
                "Projetado",
                ultimo_cobre, ultimo_aluminio, ultimo_dolar,
                calc_kg(ultimo_cobre, ultimo_dolar),
                calc_kg(ultimo_aluminio, ultimo_dolar),
            ])

    def media_col(linhas, col):
        vals = []
        for l in linhas:
            v = l[col] if col < len(l) else None
            if v not in (None, ""):
                try:
                    vals.append(float(str(v).replace(",", ".")))
                except ValueError:
                    pass
        return round(sum(vals) / len(vals), 4) if vals else None

    reais = [l for l in todas_linhas if l[2] == "Real"]
    todos = todas_linhas

    media_real = ["Média Real", "", "", media_col(reais, 3), media_col(reais, 4), media_col(reais, 5), media_col(reais, 6), media_col(reais, 7)]
    media_proj = ["Média Projetada", "", "", media_col(todos, 3), media_col(todos, 4), media_col(todos, 5), media_col(todos, 6), media_col(todos, 7)]

    novas_linhas = [CABECALHO] + todas_linhas + [[], media_real, media_proj]
    aba.clear()
    aba.update("A1", novas_linhas)

    print(f"✅ Aba atualizada: {len(reais)} reais + {len(todos)-len(reais)} projetados.")
    return todas_linhas

def atualizar_consolidado(planilha):
    """Atualiza aba Consolidado com dados de todas as abas de meses."""
    try:
        consolidado = planilha.worksheet("Consolidado")
        consolidado.clear()
    except gspread.WorksheetNotFound:
        consolidado = planilha.add_worksheet(title="Consolidado", rows=2000, cols=10)

    linhas_consolidado = [CABECALHO_CONSOLIDADO]

    for ws in planilha.worksheets():
        if ws.title in ("Consolidado",):
            continue
        # Só abas no formato Mês/Ano (ex: Jun/2026)
        if "/" not in ws.title:
            continue
        try:
            dados = ws.get_all_values()
            nome_mes = ws.title
            for linha in dados[1:]:
                if not linha or len(linha) < 3:
                    continue
                # Ignora linhas de média e vazias
                if linha[2] not in ("Real", "Projetado"):
                    continue
                linhas_consolidado.append(linha[:8] + [nome_mes])
        except Exception as e:
            print(f"  ⚠️  Erro ao ler aba {ws.title}: {e}")

    consolidado.update("A1", linhas_consolidado)
    print(f"✅ Consolidado atualizado: {len(linhas_consolidado)-1} linhas.")

def gravar_no_sheets(client, dados):
    planilha = client.open_by_key(GOOGLE_SHEET_ID)
    data_obj = dados["data_obj"]
    ano, mes = data_obj.year, data_obj.month

    aba = obter_ou_criar_aba(planilha, ano, mes)

    todas = aba.col_values(1)
    if dados["data"] not in todas:
        nova_linha = [
            dados["data"], dados["dia_semana"], "Real",
            dados["cobre_usd_t"], dados["aluminio_usd_t"], dados["dolar_brl"],
            calc_kg(dados["cobre_usd_t"], dados["dolar_brl"]),
            calc_kg(dados["aluminio_usd_t"], dados["dolar_brl"]),
        ]
        aba.append_row(nova_linha)
        print(f"✅ Linha real gravada: {nova_linha}")
    else:
        print(f"⚠️  Data {dados['data']} já existe.")

    recalcular_aba(aba, ano, mes)
    atualizar_consolidado(planilha)

def main():
    print("=" * 50)
    print(f"Iniciando coleta LME — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 50)

    dados = obter_dados_ontem()
    if not dados:
        print("Nenhum dado para gravar. Encerrando.")
        return

    print("Conectando ao Google Sheets...")
    client = conectar_google_sheets()
    gravar_no_sheets(client, dados)

    print("=" * 50)
    print("✅ Coleta concluída com sucesso!")
    print("=" * 50)

if __name__ == "__main__":
    main()
