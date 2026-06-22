"""
Script: enviar_email.py
Funcao: Gera tabela HTML com dados do mes atual e envia por e-mail
Roda: Todo dia as 7h via GitHub Actions
"""

import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
DESTINATARIOS = os.environ.get("EMAIL_DESTINATARIOS", "").split(",")
DASHBOARD_URL = "https://mzizzari-spec.github.io/automacao-lme"
MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def conectar_google_sheets():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not creds_json:
        raise ValueError("Secret GOOGLE_CREDENTIALS_JSON nao encontrado!")
    creds_dict = json.loads(creds_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def to_float(v):
    if v is None or v == "":
        return None
    try:
        s = str(v).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except:
        return None


def fmt(v, dec=2):
    n = to_float(v)
    if n is None:
        return "-"
    formatted = "{:,.{}f}".format(n, dec)
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_var(v):
    if v is None:
        return "-"
    n = to_float(v) if not isinstance(v, float) else v
    if n is None:
        return "-"
    num = n * 100
    sinal = "+" if num > 0 else ""
    cor = "#15803d" if num > 0 else "#dc2626" if num < 0 else "#6b7280"
    return '<span style="color:{};font-size:10px;">{}{:.2f}%</span>'.format(cor, sinal, num)


def calc_var(atual, anterior):
    a = to_float(atual) if not isinstance(atual, float) else atual
    b = to_float(anterior) if not isinstance(anterior, float) else anterior
    if a is None or b is None or b == 0:
        return None
    return (a - b) / b


def obter_dados_mes_atual(client):
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    nome_aba = "{}/{}".format(MESES_PT[mes_atual-1], ano_atual)

    mes_ant = mes_atual - 1
    ano_ant = ano_atual
    if mes_ant == 0:
        mes_ant = 12
        ano_ant -= 1
    nome_mes_ant = "{}/{}".format(MESES_PT[mes_ant-1], ano_ant)

    planilha = client.open_by_key(GOOGLE_SHEET_ID)
    try:
        aba = planilha.worksheet(nome_aba)
    except gspread.WorksheetNotFound:
        raise ValueError("Aba '{}' nao encontrada!".format(nome_aba))

    dados = aba.get_all_values()

    media_real_ant = None
    try:
        hist = planilha.worksheet("Historico")
        dados_hist = hist.get_all_values()
        linha_ant = next((l for l in dados_hist[1:] if len(l) > 0 and l[0] == nome_mes_ant), None)
        print("DEBUG: buscando '{}' no Historico, encontrou: {}".format(nome_mes_ant, linha_ant))
        if linha_ant:
            media_real_ant = [''] * 13
            media_real_ant[3] = linha_ant[3]
            media_real_ant[5] = linha_ant[5]
            media_real_ant[7] = linha_ant[1]
            media_real_ant[9] = linha_ant[7]
            media_real_ant[11] = linha_ant[9]
            print("Medias de {} encontradas no Historico".format(nome_mes_ant))
    except Exception as e:
        print("Historico nao encontrado: {}".format(e))

    return dados, nome_aba, media_real_ant


def gerar_html_email(dados, nome_mes, media_real_ant=None):
    hoje = datetime.now()

    linhas_dias = [l for l in dados[1:] if len(l) > 2 and l[2] in ("Real", "Projetado")]
    media_real = next((l for l in dados if len(l) > 0 and l[0] == "Media Real"), None)
    if media_real is None:
        media_real = next((l for l in dados if len(l) > 0 and "Real" in l[0] and "Proj" not in l[0]), None)
    media_proj = next((l for l in dados if len(l) > 0 and l[0] == "Media Projetada"), None)
    if media_proj is None:
        media_proj = next((l for l in dados if len(l) > 0 and "Projetada" in l[0]), None)
    reais = [l for l in linhas_dias if l[2] == "Real"]
    projetados = [l for l in linhas_dias if l[2] == "Projetado"]
    semanas = [l for l in dados[1:] if len(l) > 0 and "Semana" in l[0]]

    def card(label, val, var, cor, unidade, dec=2, width="20%"):
        var_html = fmt_var(var)
        return (
            '<td style="padding:0 4px;" width="{}">'
            '<div style="background:#fff;border:1px solid #e2e4ea;border-radius:8px;padding:12px;">'
            '<div style="font-size:10px;color:#6b7280;text-transform:uppercase;margin-bottom:4px;">{}</div>'
            '<div style="font-size:15px;font-weight:600;color:{};font-family:monospace;">{}</div>'
            '<div style="font-size:10px;color:#6b7280;margin-top:2px;">{} &nbsp; {}</div>'
            '</div></td>'
        ).format(width, label, cor, fmt(val, dec), unidade, var_html)

    def var_r(col):
        atual = to_float(media_real[col]) if media_real and len(media_real) > col else None
        ant = to_float(media_real_ant[col]) if media_real_ant and len(media_real_ant) > col else None
        return calc_var(atual, ant)

    def var_p(col):
        proj = to_float(media_proj[col]) if media_proj and len(media_proj) > col else None
        real = to_float(media_real[col]) if media_real and len(media_real) > col else None
        return calc_var(proj, real)

    cards_real = ""
    if media_real and len(media_real) > 12:
        cards_real = (
            "<tr>"
            + card("Cobre Real", media_real[3], var_r(3), "#c45e1a", "US$/t", width="20%")
            + card("Aluminio Real", media_real[5], var_r(5), "#2b6cb0", "US$/t", width="20%")
            + card("Dolar Real", media_real[7], var_r(7), "#1a7a42", "R$/US$", dec=4, width="20%")
            + card("Cobre R$/kg Real", media_real[9], var_r(9), "#c45e1a", "R$/kg", width="20%")
            + card("Aluminio R$/kg Real", media_real[11], var_r(11), "#2b6cb0", "R$/kg", width="20%")
            + "</tr>"
        )

    cards_proj = ""
    if media_proj and len(media_proj) > 12:
        cards_proj = (
            "<tr>"
            + card("Cobre Proj.", media_proj[3], var_p(3), "#c45e1a", "US$/t", width="20%")
            + card("Aluminio Proj.", media_proj[5], var_p(5), "#2b6cb0", "US$/t", width="20%")
            + card("Dolar Proj.", media_proj[7], var_p(7), "#1a7a42", "R$/US$", dec=4, width="20%")
            + card("Cobre R$/kg Proj.", media_proj[9], var_p(9), "#c45e1a", "R$/kg", width="20%")
            + card("Aluminio R$/kg Proj.", media_proj[11], var_p(11), "#2b6cb0", "R$/kg", width="20%")
            + "</tr>"
        )

    linhas_html = ""
    semana_idx = 0

    for l in dados[1:]:
        if not l or len(l) < 3:
            continue
        tipo = l[2] if len(l) > 2 else ""

        if tipo in ("Real", "Projetado"):
            if tipo == "Real":
                badge = '<span style="background:#dcfce7;color:#15803d;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600;">Real</span>'
                bg = "#ffffff"
            else:
                badge = '<span style="background:#dbeafe;color:#1d4ed8;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600;">Proj.</span>'
                bg = "#f8faff"

            def td(col, dec=2, linha=l):
                val = linha[col] if len(linha) > col else ""
                return '<td style="padding:7px 10px;font-size:11px;font-family:monospace;">{}</td>'.format(fmt(val, dec))

            def td_var(col, linha=l):
                val = linha[col] if len(linha) > col else ""
                return '<td style="padding:7px 10px;font-size:10px;text-align:center;">{}</td>'.format(fmt_var(to_float(val)))

            linhas_html += (
                '<tr style="background:{};border-bottom:1px solid #e2e4ea;">'
                '<td style="padding:7px 10px;font-size:11px;color:#6b7280;">{}</td>'
                '<td style="padding:7px 10px;font-size:11px;">{}</td>'
                '<td style="padding:7px 10px;">{}</td>'
                '{}{}{}{}{}{}{}{}{}{}'
                '</tr>'
            ).format(bg, l[0], l[1], badge,
                     td(3,0), td_var(4), td(5,0), td_var(6),
                     td(7,4), td_var(8), td(9,2), td_var(10),
                     td(11,2), td_var(12))

        elif "Semana" in l[0]:
            sem_ant = semanas[semana_idx - 1] if semana_idx > 0 else None

            def td_s(col, dec=2, linha=l):
                val = linha[col] if len(linha) > col else ""
                return '<td style="padding:6px 10px;font-size:11px;font-family:monospace;color:#6b7280;">{}</td>'.format(fmt(val, dec))

            def td_sv(col, linha=l, ant=sem_ant):
                val = to_float(linha[col]) if len(linha) > col else None
                val_ant = to_float(ant[col]) if ant and len(ant) > col else None
                return '<td style="padding:6px 10px;font-size:10px;text-align:center;">{}</td>'.format(fmt_var(calc_var(val, val_ant)))

            linhas_html += (
                '<tr style="background:#f8f9fa;border-bottom:2px solid #e2e4ea;">'
                '<td colspan="3" style="padding:6px 10px;font-size:10px;color:#9aa0b4;font-style:italic;">Media da semana</td>'
                '{}{}{}{}{}{}{}{}{}{}'
                '</tr>'
            ).format(td_s(3,0), td_sv(3), td_s(5,0), td_sv(5),
                     td_s(7,4), td_sv(7), td_s(9,2), td_sv(9),
                     td_s(11,2), td_sv(11))
            semana_idx += 1

    def tm(linha, col, dec=2):
        val = linha[col] if linha and len(linha) > col else ""
        return '<td style="padding:7px 10px;font-size:11px;font-weight:600;font-family:monospace;">{}</td>'.format(fmt(val, dec))

    def tmv_r(col):
        v = calc_var(to_float(media_real[col]) if media_real and len(media_real) > col else None,
                     to_float(media_real_ant[col]) if media_real_ant and len(media_real_ant) > col else None)
        return '<td style="padding:7px 10px;font-size:10px;text-align:center;">{}</td>'.format(fmt_var(v))

    def tmv_p(col):
        v = calc_var(to_float(media_proj[col]) if media_proj and len(media_proj) > col else None,
                     to_float(media_real[col]) if media_real and len(media_real) > col else None)
        return '<td style="padding:7px 10px;font-size:10px;text-align:center;">{}</td>'.format(fmt_var(v))

    data_hora = hoje.strftime("%d/%m/%Y %H:%M")
    data_assunto = hoje.strftime("%d/%m/%Y")

    html = (
        '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"></head>'
        '<body style="margin:0;padding:0;background:#f5f6f8;font-family:Inter,Arial,sans-serif;">'
        '<div style="max-width:960px;margin:0 auto;padding:28px 16px;">'
        '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;padding-bottom:16px;border-bottom:1px solid #e2e4ea;">'
        '<div>'
        '<h1 style="font-size:20px;font-weight:700;color:#1a1d2e;margin:0;">LME <span style="color:#1a6080;">Metais</span></h1>'
        '<p style="font-size:12px;color:#6b7280;margin:4px 0 0;">Cotacoes {} - {}</p>'
        '</div>'
        '<img src="https://mzizzari-spec.github.io/automacao-lme/GMC-logo.png" alt="Grupo Melo Cordeiro" height="48" width="160" style="height:48px;width:160px;object-fit:contain;display:block;">'
        '</div>'
        '<table width="100%" cellpadding="0" cellspacing="8" style="margin-bottom:20px;">'
        '<tr><td colspan="5" style="padding:4px 0 8px;font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;">Media Real</td></tr>'
        '{}'
        '<tr><td colspan="5" style="padding:12px 0 8px;font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;">Media Projetada</td></tr>'
        '{}'
        '</table>'
        '<div style="background:#fff;border:1px solid #e2e4ea;border-radius:10px;overflow:hidden;margin-bottom:20px;">'
        '<div style="padding:12px 16px;border-bottom:1px solid #e2e4ea;">'
        '<span style="font-size:13px;font-weight:600;">Cotacoes diarias - {}</span>'
        '</div>'
        '<table width="100%" cellpadding="0" cellspacing="0">'
        '<thead><tr style="background:#f0f1f4;">'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Data</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Dia</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Tipo</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Cobre</th>'
        '<th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Aluminio</th>'
        '<th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Dolar</th>'
        '<th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Cu R$/kg</th>'
        '<th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>'
        '<th style="padding:8px 10px;text-align:left;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Al R$/kg</th>'
        '<th style="padding:8px 10px;text-align:center;font-size:10px;color:#6b7280;text-transform:uppercase;border-bottom:1px solid #e2e4ea;">Var%</th>'
        '</tr></thead><tbody>'
        '{}'
        '<tr style="background:#f0f1f4;border-top:2px solid #e2e4ea;">'
        '<td colspan="3" style="padding:7px 10px;font-size:11px;font-weight:600;color:#15803d;">Media Real</td>'
        '{}{}{}{}{}{}{}{}{}{}'
        '</tr>'
        '<tr style="background:#f0f1f4;">'
        '<td colspan="3" style="padding:7px 10px;font-size:11px;font-weight:600;color:#1d4ed8;">Media Projetada</td>'
        '{}{}{}{}{}{}{}{}{}{}'
        '</tr>'
        '</tbody></table></div>'
        '<div style="text-align:center;padding:16px;">'
        '<a href="{}" style="background:#c45e1a;color:#fff;padding:11px 24px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:600;">Ver Dashboard Completo</a>'
        '</div>'
        '<div style="text-align:center;padding:12px 0;font-size:10px;color:#9aa0b4;">Enviado automaticamente pelo sistema LME Automacao</div>'
        '</div></body></html>'
    ).format(
        nome_mes, data_hora,
        cards_real, cards_proj,
        nome_mes,
        linhas_html,
        tm(media_real,3,0), tmv_r(3), tm(media_real,5,0), tmv_r(5),
        tm(media_real,7,4), tmv_r(7), tm(media_real,9,2), tmv_r(9),
        tm(media_real,11,2), tmv_r(11),
        tm(media_proj,3,0), tmv_p(3), tm(media_proj,5,0), tmv_p(5),
        tm(media_proj,7,4), tmv_p(7), tm(media_proj,9,2), tmv_p(9),
        tm(media_proj,11,2), tmv_p(11),
        DASHBOARD_URL
    )
    return html


def enviar_email(html, nome_mes):
    assunto = "LME Metais - {} ({})".format(nome_mes, datetime.now().strftime("%d/%m/%Y"))
    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = GMAIL_USER
    msg['To'] = GMAIL_USER
    msg['Bcc'] = ", ".join(DESTINATARIOS)
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, DESTINATARIOS, msg.as_string())
    print("E-mail enviado!")


def main():
    print("=" * 50)
    print("Iniciando envio - {}".format(datetime.now().strftime("%d/%m/%Y %H:%M")))
    print("=" * 50)
    client = conectar_google_sheets()
    dados, nome_mes, media_real_ant = obter_dados_mes_atual(client)
    html = gerar_html_email(dados, nome_mes, media_real_ant)
    enviar_email(html, nome_mes)
    print("Concluido!")


if __name__ == "__main__":
    main()
