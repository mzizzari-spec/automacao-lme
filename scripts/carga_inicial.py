"""
Script: carga_inicial.py
Função: Roda UMA VEZ SÓ para carregar os últimos 3 meses de dados históricos
        do site shockmetais.com.br/lme para o Google Sheets.
"""

import os
import json
import calendar
import re
from datetime import datetime, date, timedelta

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

FERIADOS_FIXOS = {
    (1, 1), (21, 4), (1, 5), (7, 9),
    (12, 10), (2, 11), (15, 11), (25, 12),
}

def calcular_pascoa(ano):
    a = ano % 19; b = ano // 100; c = ano % 100
    d = b // 4; e = b % 4; f = (b + 8) // 25
    g = (b - f + 1) // 3; h = (19 * a + b - d - g + 15) % 30
    i = c // 4; k = c % 4; l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(ano, mes, dia)

def feriados_moveis(ano):
    p = calcular_pascoa(ano)
    return {p - timedelta(days=48), p - timedelta(days=47),
            p - timedelta(days=2), p + timedelta(days=60)}

def eh_feriado(d):
    if (d.day, d.month) in FERIADOS_FIXOS: return True
    if date(d.year, d.month, d.day) in feriados_moveis(d.year): return True
    return False

def eh_dia_util(d):
    if d.weekday() >= 5: return False
    if eh_feriado(d): return False
    return True

def dias_uteis_do_mes(ano, mes):
    _, ultimo = calendar.monthrange(ano, mes)
    # Inclui feriados (exclui apenas fins de semana)
    return [date(ano, mes, d) for d in range(1, ultimo + 1) if date(ano, mes, d).weekday() < 5]

def nome_aba(ano, mes):
    return f"{MESES_PT[mes-1]}/{ano}"

def conectar_google_sheets():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json: raise ValueError("Secret GOOGLE_CREDENTIALS_JSON não encontrado!")
    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def limpar_numero(texto, americano=False):
    if not texto or texto.strip().lower() in ("feriado", "-", ""): return None
    try:
        s = texto.strip()
        s = s.replace(",", "") if americano else s.replace(".", "").replace(",", ".")
        return float(s)
    except ValueError:
        return None

def para_float(valor):
    if not valor: return None
    try:
        s = str(valor).strip()
        if "," in s and "." in s: s = s.replace(".", "").replace(",", ".")
        elif "," in s: s = s.replace(",", ".")
        return float(s)
    except: return None

def calc_kg(usd_t, dolar):
    if usd_t and dolar: return round((usd_t * dolar) / 1000, 4)
    return None

def calc_variacao(atual, anterior):
    if atual is None or anterior is None or anterior == 0: return None
    return round((atual - anterior) / anterior, 6)

def buscar_dados_mes(ano, mes):
    print(f"\n📅 Buscando dados de {nome_aba(ano, mes)}...")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    response = requests.get(URL_SITE, headers=headers, timeout=30)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    tabela = soup.find("table")
    if not tabela:
        print("  ⚠️  Tabela não encontrada")
        return {}

    dados_mes = {}
    dias_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    mes_en_pt = {"Jan":"Jan","Feb":"Fev","Mar":"Mar","Apr":"Abr","May":"Mai","Jun":"Jun",
                 "Jul":"Jul","Aug":"Ago","Sep":"Set","Oct":"Out","Nov":"Nov","Dec":"Dez"}

    for linha in tabela.find_all("tr"):
        colunas = linha.find_all("td")
        if not colunas or len(colunas) < 8: continue
        dia_texto = colunas[0].get_text(strip=True)
        if "média" in dia_texto.lower() or "media" in dia_texto.lower(): continue
        match = re.match(r"(\d+)/(\w+)", dia_texto)
        if not match: continue
        dia_num = int(match.group(1))
        mes_abrev_site = match.group(2)
        # Verifica se é o mês correto
        mes_nome_esperado = MESES_PT[mes-1]
        if mes_abrev_site != mes_nome_esperado: continue
        try:
            d = date(ano, mes, dia_num)
        except ValueError:
            continue
        cobre = limpar_numero(colunas[1].get_text(strip=True), americano=True)
        aluminio = limpar_numero(colunas[3].get_text(strip=True), americano=True)
        dolar = limpar_numero(colunas[7].get_text(strip=True))
        if cobre or aluminio:
            dados_mes[d] = {
                "data": d.strftime("%d/%m/%Y"),
                "dia_semana": dias_nomes[d.weekday()],
                "cobre": cobre, "aluminio": aluminio, "dolar": dolar,
            }
            print(f"  ✅ {d.strftime('%d/%m/%Y')} — Cobre: {cobre} | Alumínio: {aluminio} | Dólar: {dolar}")
    return dados_mes

def processar_mes(planilha, ano, mes, dados_reais):
    nome = nome_aba(ano, mes)
    hoje = datetime.now().date()

    try:
        aba = planilha.worksheet(nome)
        aba.clear()
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=nome, rows=200, cols=15)

    dias_uteis = dias_uteis_do_mes(ano, mes)
    dias_nomes = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]

    # Busca último valor do mês anterior
    mes_ant = mes - 1; ano_ant = ano
    if mes_ant == 0: mes_ant = 12; ano_ant -= 1
    ultimo_ant = None
    try:
        aba_ant = planilha.worksheet(nome_aba(ano_ant, mes_ant))
        dados_ant = aba_ant.get_all_values()
        for linha in reversed(dados_ant):
            if len(linha) > 2 and linha[2] == "Real":
                ultimo_ant = {"cobre": para_float(linha[3]), "aluminio": para_float(linha[5]), "dolar": para_float(linha[7])}
                break
    except: pass

    # Último valor real disponível para projeção
    ultimo_real = None
    hoje_date = hoje.date() if hasattr(hoje, 'date') else hoje
    # Remove o dia atual dos dados reais - deve ser Projetado
    dados_reais = {k: v for k, v in dados_reais.items() if k != hoje_date}
    for d in sorted(dados_reais.keys(), reverse=True):
        ultimo_real = dados_reais[d]
        break

    todas_linhas = [CABECALHO]
    semana_atual = []
    valor_anterior = ultimo_ant

    for d in dias_uteis:
        data_str = d.strftime("%d/%m/%Y")
        if d in dados_reais:
            v = dados_reais[d]
            tipo = "Real"
            cobre = v["cobre"]; aluminio = v["aluminio"]
            dolar = v["dolar"]
            # Se dolar for None usa o último conhecido
            if dolar is None and 'ultimo_dolar' in dir():
                dolar = ultimo_dolar
            if dolar is not None:
                ultimo_dolar = dolar
        elif ultimo_real and d > hoje:
            tipo = "Projetado"
            cobre = ultimo_real["cobre"]; aluminio = ultimo_real["aluminio"]; dolar = ultimo_real["dolar"]
        else:
            continue

        cobre_kg = calc_kg(cobre, dolar)
        aluminio_kg = calc_kg(aluminio, dolar)
        var_cobre = calc_variacao(cobre, valor_anterior["cobre"] if valor_anterior else None)
        var_al = calc_variacao(aluminio, valor_anterior["aluminio"] if valor_anterior else None)
        var_dol = calc_variacao(dolar, valor_anterior["dolar"] if valor_anterior else None)
        prev_cobre_kg = calc_kg(valor_anterior["cobre"], valor_anterior["dolar"]) if valor_anterior else None
        prev_al_kg = calc_kg(valor_anterior["aluminio"], valor_anterior["dolar"]) if valor_anterior else None
        var_cobre_kg = calc_variacao(cobre_kg, prev_cobre_kg)
        var_al_kg = calc_variacao(aluminio_kg, prev_al_kg)

        linha = [data_str, dias_nomes[d.weekday()] if tipo == "Projetado" else v.get("dia_semana", dias_nomes[d.weekday()]),
                 tipo, cobre, var_cobre, aluminio, var_al, dolar, var_dol,
                 cobre_kg, var_cobre_kg, aluminio_kg, var_al_kg]
        todas_linhas.append(linha)
        semana_atual.append(linha)

        # Fim de semana?
        proximo = next((dd for dd in dias_uteis if dd > d), None)
        fim_semana = (d.weekday() == 4 or proximo is None or proximo.isocalendar()[1] != d.isocalendar()[1])
        if fim_semana and semana_atual:
            def med(col):
                vals = [para_float(l[col]) for l in semana_atual if len(l) > col and para_float(l[col]) is not None]
                return round(sum(vals)/len(vals), 4) if vals else None
            m3=med(3); m5=med(5); m7=med(7); m9=med(9); m11=med(11)
            media_ant = next((l for l in reversed(todas_linhas) if l and l[0] == "Média Semana"), None)
            var3 = calc_variacao(m3, para_float(media_ant[3])) if media_ant else None
            var5 = calc_variacao(m5, para_float(media_ant[5])) if media_ant else None
            var7 = calc_variacao(m7, para_float(media_ant[7])) if media_ant else None
            var9 = calc_variacao(m9, para_float(media_ant[9])) if media_ant else None
            var11 = calc_variacao(m11, para_float(media_ant[11])) if media_ant else None
            todas_linhas.append(["Média Semana","","",m3,var3,m5,var5,m7,var7,m9,var9,m11,var11])
            semana_atual = []

        valor_anterior = {"cobre": cobre, "aluminio": aluminio, "dolar": dolar}

    # Médias do mês
    reais = [l for l in todas_linhas[1:] if isinstance(l[2], str) and l[2] == "Real"]
    todos = [l for l in todas_linhas[1:] if isinstance(l[2], str) and l[2] in ("Real","Projetado")]

    def med_mes(linhas, col):
        vals = [para_float(l[col]) for l in linhas if len(l) > col and para_float(l[col]) is not None]
        return round(sum(vals)/len(vals), 4) if vals else None

    todas_linhas.append([])
    todas_linhas.append(["Média Real","","",med_mes(reais,3),med_mes(reais,4),med_mes(reais,5),med_mes(reais,6),med_mes(reais,7),med_mes(reais,8),med_mes(reais,9),med_mes(reais,10),med_mes(reais,11),med_mes(reais,12)])
    todas_linhas.append(["Média Projetada","","",med_mes(todos,3),med_mes(todos,4),med_mes(todos,5),med_mes(todos,6),med_mes(todos,7),med_mes(todos,8),med_mes(todos,9),med_mes(todos,10),med_mes(todos,11),med_mes(todos,12)])

    aba.update(values=todas_linhas, range_name="A1")
    print(f"  ✅ '{nome}': {len(reais)} reais + {len(todos)-len(reais)} projetados")
    return [l for l in todas_linhas[1:] if len(l) > 2 and isinstance(l[2], str) and l[2] in ("Real","Projetado")], nome

def atualizar_consolidado(planilha, todos_os_dados):
    try:
        consolidado = planilha.worksheet("Consolidado")
        consolidado.clear()
    except gspread.WorksheetNotFound:
        consolidado = planilha.add_worksheet(title="Consolidado", rows=2000, cols=15)

    linhas = [CABECALHO_CONSOLIDADO]
    for nome_mes, dados in todos_os_dados:
        for l in dados:
            while len(l) < 13: l.append("")
            linhas.append(l[:13] + [nome_mes])

    consolidado.update(values=linhas, range_name="A1")
    print(f"\n✅ Consolidado: {len(linhas)-1} linhas.")

def main():
    print("="*55)
    print("CARGA INICIAL — Últimos 3 meses")
    print("="*55)

    hoje = datetime.now()
    client = conectar_google_sheets()
    planilha = client.open_by_key(GOOGLE_SHEET_ID)

    todos_os_dados = []
    for i in range(2, -1, -1):
        mes_alvo = hoje.month - i
        ano_alvo = hoje.year
        while mes_alvo <= 0: mes_alvo += 12; ano_alvo -= 1

        dados_reais = buscar_dados_mes(ano_alvo, mes_alvo)
        if not dados_reais and i > 0:
            print("  ⚠️  Sem dados, pulando.")
            continue

        linhas, nome = processar_mes(planilha, ano_alvo, mes_alvo, dados_reais)
        todos_os_dados.append((nome, linhas))

    atualizar_consolidado(planilha, todos_os_dados)
    print("\n"+"="*55)
    print("✅ Carga inicial concluída!")
    print("="*55)

if __name__ == "__main__":
    main()
