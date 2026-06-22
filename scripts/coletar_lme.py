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
            # Nunca pega o dia atual como Real
            if ontem.date() == datetime.now().date():
                print(f"⚠️  Dia '{dia_site}' é hoje — ignorando para manter como Projetado.")
                return None
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
    # Inclui linhas inseridas manualmente (mesmo com valores vazios)
    for r in todos_registros[1:]:
        if r and len(r) > 2 and r[2] == "Real" and r[0] not in datas_reais:
            datas_reais[r[0]] = r
    dias_uteis = dias_uteis_do_mes(ano, mes)
    dias_semana_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]
    hoje = datetime.now().date()

    # Busca último valor do mês anterior para variação do primeiro dia
    ultimo_mes_ant = obter_ultimo_valor_mes_anterior(planilha, ano, mes)

    # Monta lista de valores por data para calcular variações
    ultimo_dolar_conhecido = ultimo_mes_ant["dolar"] if ultimo_mes_ant else None
    valores_por_data = {}
    for d in dias_uteis:
        data_str = d.strftime("%d/%m/%Y")
        if data_str in datas_reais:
            r = datas_reais[data_str]
            dolar = para_float(r[7])
            # Se dolar for None, usa o último conhecido
            if dolar is None:
                dolar = ultimo_dolar_conhecido
            else:
                ultimo_dolar_conhecido = dolar
            cobre = para_float(r[3])
            aluminio = para_float(r[5])
            valores_por_data[data_str] = {
                "cobre": cobre,
                "aluminio": aluminio,
                "dolar": dolar,
                "cobre_kg": calc_kg(cobre, dolar),
                "aluminio_kg": calc_kg(aluminio, dolar),
                "tipo": "Real",
                "dia_semana": r[1],
            }
        elif d >= hoje:
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
            if d < hoje:
                # Dia passado sem dados — linha vazia para preenchimento manual
                dias_semana_nomes_local = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta"]
                linha_vazia = [data_str, dias_semana_nomes_local[d.weekday()], "Real",
                               "", "", "", "", "", "", "", "", "", ""]
                todas_linhas.append(linha_vazia)
                semana_atual.append(linha_vazia)
                proximo_dia_util = next((dd for dd in dias_uteis if dd > d), None)
                num_semana = d.isocalendar()[1]
                fim_semana = (d.weekday() == 4 or proximo_dia_util is None or
                             proximo_dia_util.isocalendar()[1] != num_semana)
                if fim_semana and semana_atual:
                    media_anterior = next(
                        (l for l in reversed(todas_linhas) if l and l[0] == "Media Semana"), None)
                    media_linha = calcular_media_semana(semana_atual, media_anterior)
                    todas_linhas.append(media_linha)
                    semana_atual = []
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
            # Busca média da semana anterior nas linhas já calculadas
            media_anterior = next(
                (l for l in reversed(todas_linhas) if l and l[0] == "Média Semana"),
                None
            )
            media_linha = calcular_media_semana(semana_atual, media_anterior)
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
    aba.update(values=todas_linhas, range_name="A1")
    print(f"✅ Aba atualizada: {len(reais)} reais + {len(todos)-len(reais)} projetados.")
    return [l for l in todas_linhas[1:] if len(l) > 2 and isinstance(l[2], str) and l[2] in ("Real", "Projetado")]


def calcular_media_semana(linhas, media_semana_anterior=None):
    """Calcula média de uma semana com variação em relação à semana anterior."""
    def med(col):
        vals = [para_float(l[col]) for l in linhas if len(l) > col and para_float(l[col]) is not None]
        return round(sum(vals) / len(vals), 4) if vals else None

    m3 = med(3); m5 = med(5); m7 = med(7); m9 = med(9); m11 = med(11)

    if media_semana_anterior:
        var3 = calc_variacao(m3, para_float(media_semana_anterior[3]))
        var5 = calc_variacao(m5, para_float(media_semana_anterior[5]))
        var7 = calc_variacao(m7, para_float(media_semana_anterior[7]))
        var9 = calc_variacao(m9, para_float(media_semana_anterior[9]))
        var11 = calc_variacao(m11, para_float(media_semana_anterior[11]))
    else:
        var3 = var5 = var7 = var9 = var11 = None

    return ["Média Semana", "", "", m3, var3, m5, var5, m7, var7, m9, var9, m11, var11]


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
    aba.update(values=linhas_resumo, range_name=f"A{linha_inicio}")
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

    consolidado.update(values=linhas_consolidado, range_name="A1")
    print(f"✅ Consolidado atualizado: {len(linhas_consolidado)-1} linhas.")


def atualizar_historico(planilha, ano, mes):
    """Adiciona o mês atual ao Historico se ainda não existir."""
    nome_mes = nome_aba(ano, mes)
    try:
        aba_mes = planilha.worksheet(nome_mes)
        dados_mes = aba_mes.get_all_values()
    except:
        return

    # Busca linha Média Real da aba do mês
    media_real = next((l for l in dados_mes if l and l[0] == "Média Real"), None)
    if not media_real or len(media_real) < 12:
        return

    # Monta linha do resumo
    label = nome_mes
    nova_linha = [
        label,
        para_float(media_real[7]), None,  # dolar, var
        para_float(media_real[3]), None,  # cobre, var
        para_float(media_real[5]), None,  # aluminio, var
        para_float(media_real[9]), None,  # cobre_kg, var
        para_float(media_real[11]), None, # aluminio_kg, var
    ]

    try:
        hist = planilha.worksheet("Historico")
    except gspread.WorksheetNotFound:
        print("⚠️  Aba Historico não encontrada. Rode importar_historico primeiro.")
        return

    todos = hist.get_all_values()
    labels_existentes = [l[0] for l in todos if l]

    if label in labels_existentes:
        # Atualiza linha existente
        idx = labels_existentes.index(label)
        # Calcula variação em relação ao mês anterior
        if idx > 1:
            anterior = todos[idx - 1]
            nova_linha[2] = calc_variacao(nova_linha[1], para_float(anterior[1]))  # var dolar
            nova_linha[4] = calc_variacao(nova_linha[3], para_float(anterior[3]))  # var cobre
            nova_linha[6] = calc_variacao(nova_linha[5], para_float(anterior[5]))  # var aluminio
            nova_linha[8] = calc_variacao(nova_linha[7], para_float(anterior[7]))  # var cobre_kg
            nova_linha[10] = calc_variacao(nova_linha[9], para_float(anterior[9])) # var aluminio_kg
        hist.update(values=[nova_linha], range_name=f"A{idx + 1}")
        print(f"✅ Historico: linha '{label}' atualizada.")
    else:
        # Adiciona nova linha
        if len(todos) > 1:
            anterior = todos[-1]
            nova_linha[2] = calc_variacao(nova_linha[1], para_float(anterior[1]))
            nova_linha[4] = calc_variacao(nova_linha[3], para_float(anterior[3]))
            nova_linha[6] = calc_variacao(nova_linha[5], para_float(anterior[5]))
            nova_linha[8] = calc_variacao(nova_linha[7], para_float(anterior[7]))
            nova_linha[10] = calc_variacao(nova_linha[9], para_float(anterior[9]))
        hist.append_row(nova_linha)
        print(f"✅ Historico: linha '{label}' adicionada.")


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

    # Atualiza historico
    atualizar_historico(planilha, ano, mes)


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
