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
    "Cobre (US$/t)", "Var. Cobre",
    "Alumínio (US$/t)", "Var. Alumínio",
    "Dólar (R$/US$)", "Var. Dólar",
    "Cobre (R$/kg)", "Var. Cobre/kg",
    "Alumínio (R$/kg)", "Var. Alumínio/kg",
]

CABECALHO_CONSOLIDADO = CABECALHO + ["Mês"]

CABECALHO_RESUMO = [
    "Mês", 
    "Dólar", "Var. Dólar",
    "Cobre (US$/t)", "Var. Cobre",
    "Alumínio (US$/t)", "Var. Alumínio",
    "Cobre (R$/kg)", "Var. Cobre/kg",
    "Alumínio (R$/kg)", "Var. Alumínio/kg",
]

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
            s = s.replace(",", "")
        else:
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except ValueError:
        return None

def para_float(valor):
    if not valor:
        return None
    try:
        s = str(valor).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except:
        return None

def calc_kg(usd_t, dolar):
    if usd_t and dolar:
        return round((usd_t * dolar) / 1000, 4)
    return None

def calc_variacao(atual, anterior):
    if atual is None or anterior is None or anterior == 0:
        return None
    return round((atual - anterior) / anterior, 6)

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
        aba = planilha.add_worksheet(title=nome, rows=200, cols=15)
        print(f"Aba '{nome}' criada.")
    return aba

def obter_ultimo_valor_mes_anterior(planilha, ano, mes):
    """Busca o último valor real do mês anterior para calcular variação do primeiro dia."""
    mes_ant = mes - 1
    ano_ant = ano
    if mes_ant == 0:
        mes_ant = 12
        ano_ant -= 1

    nome = nome_aba(ano_ant, mes_ant)
    try:
        aba = planilha.worksheet(nome)
        dados = aba.get_all_values()
        # Busca última linha Real
        for linha in reversed(dados):
            if len(linha) > 2 and linha[2] == "Real":
                return {
                    "cobre": para_float(linha[3]),
                    "aluminio": para_float(linha[5]),
                    "dolar": para_float(linha[7]),
                    "cobre_kg": para_float(linha[9]),
                    "aluminio_kg": para_float(linha[11]),
                }
    except:
        pass
    return None

def recalcular_aba(planilha, aba, ano, mes):
    """Recalcula toda a aba com variações, médias semanais e resumo mensal."""
    todos_registros = aba.get_all_values()
    if not todos_registros:
        return []

    # Pega só linhas reais (sem cabeçalho, médias etc)
    linhas_reais = [r for r in todos_registros[1:] if r and len(r) > 2 and r[2] == "Real"]
    if not linhas_reais:
        return []

    # Último valor real para projeções
    ultimo = linhas_reais[-1]
    ultimo_cobre = para_float(ultimo[3])
    ultimo_aluminio = para_float(ultimo[5])
    ultimo_dolar = para_float(ultimo[7])

    datas_reais = {r[0]: r for r in linhas_reais}
    dias_uteis = dias_uteis_do_mes(ano, mes)
    dias_semana_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
    hoje = datetime.now().date()

    # Busca último valor do mês anterior para variação do primeiro dia
    ultimo_mes_ant = obter_ultimo_valor_mes_anterior(planilha, ano, mes)

    # Monta lista de valores por data para calcular variações
    valores_por_data = {}
    for d in dias_uteis:
        data_str = d.strftime("%d/%m/%Y")
        if data_str in datas_reais:
            r = datas_reais[data_str]
            valores_por_data[data_str] = {
                "cobre": para_float(r[3]),
                "aluminio": para_float(r[5]),
                "dolar": para_float(r[7]),
                "cobre_kg": para_float(r[9]) if len(r) > 9 else None,
                "aluminio_kg": para_float(r[11]) if len(r) > 11 else None,
                "tipo": "Real",
                "dia_semana": r[1],
            }
        elif d > hoje:
            cobre_kg = calc_kg(ultimo_cobre, ultimo_dolar)
            aluminio_kg = calc_kg(ultimo_aluminio, ultimo_dolar)
            valores_por_data[data_str] = {
                "cobre": ultimo_cobre,
                "aluminio": ultimo_aluminio,
                "dolar": ultimo_dolar,
                "cobre_kg": cobre_kg,
                "aluminio_kg": aluminio_kg,
                "tipo": "Projetado",
                "dia_semana": dias_semana_nomes[d.weekday()],
            }

    # Monta linhas com variações e médias semanais
    todas_linhas = [CABECALHO]
    semana_atual = []
    num_semana_anterior = None
    valor_anterior = ultimo_mes_ant

    for d in dias_uteis:
        data_str = d.strftime("%d/%m/%Y")
        if data_str not in valores_por_data:
            continue

        v = valores_por_data[data_str]

        # Calcula variações
        var_cobre = calc_variacao(v["cobre"], valor_anterior["cobre"] if valor_anterior else None)
        var_al = calc_variacao(v["aluminio"], valor_anterior["aluminio"] if valor_anterior else None)
        var_dol = calc_variacao(v["dolar"], valor_anterior["dolar"] if valor_anterior else None)

        cobre_kg = v["cobre_kg"] or calc_kg(v["cobre"], v["dolar"])
        aluminio_kg = v["aluminio_kg"] or calc_kg(v["aluminio"], v["dolar"])
        prev_cobre_kg = calc_kg(valor_anterior["cobre"], valor_anterior["dolar"]) if valor_anterior else None
        prev_al_kg = calc_kg(valor_anterior["aluminio"], valor_anterior["dolar"]) if valor_anterior else None

        var_cobre_kg = calc_variacao(cobre_kg, prev_cobre_kg)
        var_al_kg = calc_variacao(aluminio_kg, prev_al_kg)

        linha = [
            data_str, v["dia_semana"], v["tipo"],
            v["cobre"], var_cobre,
            v["aluminio"], var_al,
            v["dolar"], var_dol,
            cobre_kg, var_cobre_kg,
            aluminio_kg, var_al_kg,
        ]
        todas_linhas.append(linha)
        semana_atual.append(linha)

        # Verifica se acabou a semana (sexta ou último dia útil do mês)
        num_semana = d.isocalendar()[1]
        proximo_dia_util = next((dd for dd in dias_uteis if dd > d), None)
        fim_semana = (d.weekday() == 4 or proximo_dia_util is None or
                     proximo_dia_util.isocalendar()[1] != num_semana)

        if fim_semana and semana_atual:
            media_linha = calcular_media_semana(semana_atual)
            todas_linhas.append(media_linha)
            semana_atual = []

        valor_anterior = {
            "cobre": v["cobre"],
            "aluminio": v["aluminio"],
            "dolar": v["dolar"],
        }

    # Médias do mês
    reais = [l for l in todas_linhas[1:] if isinstance(l[2], str) and l[2] == "Real"]
    todos = [l for l in todas_linhas[1:] if isinstance(l[2], str) and l[2] in ("Real", "Projetado")]

    todas_linhas.append([])
    todas_linhas.append(calcular_media_mes(reais, "Média Real"))
    todas_linhas.append(calcular_media_mes(todos, "Média Projetada"))

    # Atualiza aba
    aba.clear()
    aba.update("A1", todas_linhas)
    print(f"✅ Aba atualizada: {len(reais)} reais + {len(todos)-len(reais)} projetados.")
    return [l for l in todas_linhas[1:] if isinstance(l[2], str) and l[2] in ("Real", "Projetado")]


def calcular_media_semana(linhas):
    """Calcula média de uma semana."""
    def med(col):
        vals = [para_float(l[col]) for l in linhas if len(l) > col and para_float(l[col]) is not None]
        return round(sum(vals) / len(vals), 4) if vals else None

    return ["Média Semana", "", "", med(3), med(4), med(5), med(6), med(7), med(8), med(9), med(10), med(11), med(12)]


def calcular_media_mes(linhas, label):
    """Calcula média do mês."""
    def med(col):
        vals = [para_float(l[col]) for l in linhas if len(l) > col and para_float(l[col]) is not None]
        return round(sum(vals) / len(vals), 4) if vals else None

    return [label, "", "", med(3), med(4), med(5), med(6), med(7), med(8), med(9), med(10), med(11), med(12)]


def atualizar_resumo_mensal(planilha, aba, ano, mes):
    """Atualiza o resumo mensal abaixo dos dados na aba."""
    # Busca todas as abas de meses
    todos_dados = aba.get_all_values()
    num_linhas = len(todos_dados)

    # Coleta médias reais de cada aba de mês disponível
    resumo = []
    for ws in sorted(planilha.worksheets(), key=lambda x: x.title):
        if "/" not in ws.title or ws.title == "Consolidado":
            continue
        try:
            partes = ws.title.split("/")
            mes_nome = partes[0]
            ano_ws = int(partes[1])
            mes_ws = MESES_PT.index(mes_nome) + 1
        except:
            continue

        try:
            dados_ws = ws.get_all_values()
            # Busca linha Média Real
            for linha in dados_ws:
                if linha and linha[0] == "Média Real":
                    resumo.append({
                        "label": ws.title,
                        "dolar": para_float(linha[7]) if len(linha) > 7 else None,
                        "cobre": para_float(linha[3]) if len(linha) > 3 else None,
                        "aluminio": para_float(linha[5]) if len(linha) > 5 else None,
                        "cobre_kg": para_float(linha[9]) if len(linha) > 9 else None,
                        "aluminio_kg": para_float(linha[11]) if len(linha) > 11 else None,
                    })
                    break
        except:
            continue

    if not resumo:
        return

    # Monta linhas do resumo com variações
    linhas_resumo = [[], CABECALHO_RESUMO]
    anterior = None
    for r in resumo:
        var_dol = calc_variacao(r["dolar"], anterior["dolar"] if anterior else None)
        var_cobre = calc_variacao(r["cobre"], anterior["cobre"] if anterior else None)
        var_al = calc_variacao(r["aluminio"], anterior["aluminio"] if anterior else None)
        var_cobre_kg = calc_variacao(r["cobre_kg"], anterior["cobre_kg"] if anterior else None)
        var_al_kg = calc_variacao(r["aluminio_kg"], anterior["aluminio_kg"] if anterior else None)

        linhas_resumo.append([
            r["label"],
            r["dolar"], var_dol,
            r["cobre"], var_cobre,
            r["aluminio"], var_al,
            r["cobre_kg"], var_cobre_kg,
            r["aluminio_kg"], var_al_kg,
        ])
        anterior = r

    # Grava abaixo dos dados existentes
    linha_inicio = num_linhas + 2
    aba.update(f"A{linha_inicio}", linhas_resumo)
    print(f"✅ Resumo mensal atualizado com {len(resumo)} meses.")


def atualizar_consolidado(planilha):
    """Atualiza aba Consolidado com dados de todas as abas de meses."""
    try:
        consolidado = planilha.worksheet("Consolidado")
        consolidado.clear()
    except gspread.WorksheetNotFound:
        consolidado = planilha.add_worksheet(title="Consolidado", rows=2000, cols=15)

    linhas_consolidado = [CABECALHO_CONSOLIDADO]

    for ws in planilha.worksheets():
        if ws.title == "Consolidado" or "/" not in ws.title:
            continue
        try:
            dados = ws.get_all_values()
            nome_mes = ws.title
            for linha in dados[1:]:
                if not linha or len(linha) < 3:
                    continue
                if linha[2] not in ("Real", "Projetado"):
                    continue
                # Garante 13 colunas
                while len(linha) < 13:
                    linha.append("")
                linhas_consolidado.append(linha[:13] + [nome_mes])
        except Exception as e:
            print(f"  ⚠️  Erro ao ler aba {ws.title}: {e}")

    consolidado.update("A1", linhas_consolidado)
    print(f"✅ Consolidado atualizado: {len(linhas_consolidado)-1} linhas.")


def gravar_no_sheets(client, dados):
    planilha = client.open_by_key(GOOGLE_SHEET_ID)
    data_obj = dados["data_obj"]
    ano, mes = data_obj.year, data_obj.month

    aba = obter_ou_criar_aba(planilha, ano, mes)

    # Verifica se data já existe
    todas = aba.col_values(1)
    if dados["data"] not in todas:
        cobre_kg = calc_kg(dados["cobre_usd_t"], dados["dolar_brl"])
        aluminio_kg = calc_kg(dados["aluminio_usd_t"], dados["dolar_brl"])

        nova_linha = [
            dados["data"], dados["dia_semana"], "Real",
            dados["cobre_usd_t"], None,
            dados["aluminio_usd_t"], None,
            dados["dolar_brl"], None,
            cobre_kg, None,
            aluminio_kg, None,
        ]
        aba.append_row(nova_linha)
        print(f"✅ Linha real gravada.")
    else:
        print(f"⚠️  Data {dados['data']} já existe.")

    # Recalcula tudo
    linhas_dias = recalcular_aba(planilha, aba, ano, mes)

    # Atualiza resumo mensal
    atualizar_resumo_mensal(planilha, aba, ano, mes)

    # Atualiza consolidado
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
