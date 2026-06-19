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


# ─────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
URL_SITE = "https://shockmetais.com.br/lme"

CABECALHO = [
    "Data", "Dia da Semana", "Tipo",
    "Cobre (US$/t)", "Alumínio (US$/t)", "Dólar (R$/US$)",
    "Cobre (R$/kg)", "Alumínio (R$/kg)",
]

# Feriados nacionais fixos (dia, mês)
FERIADOS_FIXOS = {
    (1, 1),   # Confraternização Universal
    (21, 4),  # Tiradentes
    (1, 5),   # Dia do Trabalho
    (7, 9),   # Independência
    (12, 10), # Nossa Senhora Aparecida
    (2, 11),  # Finados
    (15, 11), # Proclamação da República
    (25, 12), # Natal
}

# Feriados móveis por ano (calculados manualmente para os próximos anos)
# Carnaval = 47 dias antes da Páscoa, Sexta-feira Santa = 2 dias antes, Corpus Christi = 60 dias depois
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
    carnaval1 = pascoa - timedelta(days=48)
    carnaval2 = pascoa - timedelta(days=47)
    sexta_santa = pascoa - timedelta(days=2)
    corpus = pascoa + timedelta(days=60)
    return {carnaval1, carnaval2, sexta_santa, corpus}

def eh_feriado_nacional(d):
    if (d.day, d.month) in FERIADOS_FIXOS:
        return True
    if date(d.year, d.month, d.day) in feriados_moveis(d.year):
        return True
    return False

def eh_dia_util(d):
    if d.weekday() >= 5:  # sábado=5, domingo=6
        return False
    if eh_feriado_nacional(d):
        return False
    return True

def dias_uteis_do_mes(ano, mes):
    """Retorna lista de todos os dias úteis do mês."""
    _, ultimo_dia = calendar.monthrange(ano, mes)
    dias = []
    for d in range(1, ultimo_dia + 1):
        dt = date(ano, mes, d)
        if eh_dia_util(dt):
            dias.append(dt)
    return dias


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
    client = gspread.authorize(creds)
    return client


def limpar_numero(texto):
    if not texto or texto.strip().lower() in ("feriado", "-", ""):
        return None
    texto = texto.strip().replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
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

    for linha in tabela.find_all("tr"):
        colunas = linha.find_all("td")
        if not colunas:
            continue
        dia_celula = colunas[0].get_text(strip=True)
        if dia_site in dia_celula:
            cobre = limpar_numero(colunas[1].get_text(strip=True)) if len(colunas) > 1 else None
            aluminio = limpar_numero(colunas[3].get_text(strip=True)) if len(colunas) > 3 else None
            dolar = limpar_numero(colunas[7].get_text(strip=True)) if len(colunas) > 7 else None

            dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
            return {
                "data": ontem.strftime("%d/%m/%Y"),
                "data_obj": ontem.date(),
                "dia_semana": dias_semana[ontem.weekday()],
                "cobre_usd_t": cobre,
                "aluminio_usd_t": aluminio,
                "dolar_brl": dolar,
            }

    print(f"⚠️  Dia '{dia_site}' não encontrado (feriado ou fim de semana).")
    return None


def nome_aba(ano, mes):
    meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
             "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
    return f"{meses[mes-1]}/{ano}"


def obter_ou_criar_aba(planilha, ano, mes):
    nome = nome_aba(ano, mes)
    try:
        aba = planilha.worksheet(nome)
        print(f"Aba '{nome}' encontrada.")
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome, rows=200, cols=10)
        aba.append_row(CABECALHO)
        print(f"Aba '{nome}' criada.")
    return aba


def calcular_linhas_projecao(aba, dados_novos, ano, mes):
    """
    Recalcula todas as linhas projetadas e as médias no final da aba.
    """
    todos_registros = aba.get_all_values()
    if not todos_registros:
        return

    # Separa cabeçalho e linhas de dados reais
    cabecalho_linha = todos_registros[0]
    linhas_reais = [r for r in todos_registros[1:] if r and r[2] == "Real"]

    if not linhas_reais:
        return

    # Último valor real
    ultimo = linhas_reais[-1]
    ultimo_cobre = float(ultimo[3]) if ultimo[3] else None
    ultimo_aluminio = float(ultimo[4]) if ultimo[4] else None
    ultimo_dolar = float(ultimo[5]) if ultimo[5] else None

    def calc_kg(usd_t, dolar):
        if usd_t and dolar:
            return round((usd_t * dolar) / 1000, 4)
        return None

    # Datas reais já registradas
    datas_reais = set(r[0] for r in linhas_reais)

    # Todos os dias úteis do mês
    dias_uteis = dias_uteis_do_mes(ano, mes)
    dias_semana_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]

    hoje = datetime.now().date()

    # Monta todas as linhas: reais + projetadas
    todas_linhas = []
    for d in dias_uteis:
        data_str = d.strftime("%d/%m/%Y")
        dia_semana = dias_semana_nomes[d.weekday()]

        if data_str in datas_reais:
            # Linha real já existe — pega do registro
            linha_real = next(r for r in linhas_reais if r[0] == data_str)
            todas_linhas.append(linha_real[:8])
        elif d > hoje:
            # Dia futuro — projeção
            cobre_kg = calc_kg(ultimo_cobre, ultimo_dolar)
            aluminio_kg = calc_kg(ultimo_aluminio, ultimo_dolar)
            todas_linhas.append([
                data_str, dia_semana, "Projetado",
                ultimo_cobre, ultimo_aluminio, ultimo_dolar,
                cobre_kg, aluminio_kg
            ])

    # Calcula médias
    reais = [l for l in todas_linhas if l[2] == "Real"]
    projetados = [l for l in todas_linhas if l[2] == "Projetado"]
    todos = reais + projetados

    def media_col(linhas, col):
        vals = [float(l[col]) for l in linhas if l[col] not in (None, "")]
        return round(sum(vals) / len(vals), 4) if vals else None

    media_real = [
        "Média Real", "", "",
        media_col(reais, 3), media_col(reais, 4), media_col(reais, 5),
        media_col(reais, 6), media_col(reais, 7)
    ]
    media_proj = [
        "Média Projetada", "", "",
        media_col(todos, 3), media_col(todos, 4), media_col(todos, 5),
        media_col(todos, 6), media_col(todos, 7)
    ]

    # Reescreve a aba inteira
    novas_linhas = [CABECALHO] + todas_linhas + [[], media_real, media_proj]
    aba.clear()
    aba.update("A1", novas_linhas)
    print(f"✅ Aba atualizada: {len(reais)} dias reais + {len(projetados)} projetados.")


def gravar_no_sheets(client, dados):
    planilha = client.open_by_key(GOOGLE_SHEET_ID)

    data_obj = dados["data_obj"]
    ano, mes = data_obj.year, data_obj.month

    aba = obter_ou_criar_aba(planilha, ano, mes)

    # Verifica se data já existe
    todas = aba.col_values(1)
    if dados["data"] in todas:
        print(f"⚠️  Data {dados['data']} já existe. Atualizando projeções mesmo assim.")
    else:
        # Adiciona linha real
        def calc_kg(usd_t, dolar):
            if usd_t and dolar:
                return round((usd_t * dolar) / 1000, 4)
            return None

        nova_linha = [
            dados["data"],
            dados["dia_semana"],
            "Real",
            dados["cobre_usd_t"],
            dados["aluminio_usd_t"],
            dados["dolar_brl"],
            calc_kg(dados["cobre_usd_t"], dados["dolar_brl"]),
            calc_kg(dados["aluminio_usd_t"], dados["dolar_brl"]),
        ]
        aba.append_row(nova_linha)
        print(f"✅ Linha real gravada: {nova_linha}")

    # Recalcula projeções e médias
    calcular_linhas_projecao(aba, dados, ano, mes)


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
