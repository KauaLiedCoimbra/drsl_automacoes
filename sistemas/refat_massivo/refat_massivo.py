import os
import time
import math
import pandas as pd
import win32com.client as win32
from win32com.client import constants
import string
import utils as u

# === CONFIGURAÇÕES PADRÃO ===
PASTA_DOWNLOAD_PADRAO = r"C:\Users\2038860\OneDrive - CPFL Energia S A\projetos\refat massivo\relatorios_temp"

# ---------------------------
# Conexão SAP
# ---------------------------
def conectar_sap():
    SapGuiAuto = win32.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session

def abrir_transacao(session, periodo, p_file):
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFAT0657"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radP_ACES2").select()
    session.findById("wnd[0]/usr/txtS_BPER-LOW").text = periodo
    session.findById("wnd[0]/usr/ctxtP_FILE").text = p_file

# ---------------------------
# Processamento de lotes
# ---------------------------
def processar_lote(logs_widget, session, ws, start_row, end_row, indice, pasta_download=PASTA_DOWNLOAD_PADRAO, coluna="Instalação"):
    cabecalhos = [ws.Cells(1, col).Value for col in range(1, ws.UsedRange.Columns.Count + 1)]
    try:
        col_index = cabecalhos.index(coluna) + 1  # +1 porque Excel é 1-based
    except ValueError:
        raise ValueError(f"Coluna '{coluna}' não encontrada na planilha.")

    # Converter número da coluna em letra (1 → A, 2 → B, 3 → C...)
    col_letter = string.ascii_uppercase[col_index - 1]

    # Abrir múltipla seleção
    session.findById("wnd[0]/usr/btn%_S_ANLG_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()  # Limpar conteúdo antigo

    # Copiar intervalo do Excel
    intervalo = ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}")
    intervalo.Copy()

    # Colar no SAP
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Confirmar seleção

    # Executar consulta
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Exportar para Excel
    nome_temp = f"relatorio_temp_{indice}.XLS"
    session.findById("wnd[0]").sendVKey(45)
    session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta_download
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_temp
    session.findById("wnd[0]").sendVKey(0)

    # Esperar arquivo ser criado
    caminho_temp = os.path.join(pasta_download, nome_temp)
    while True:
        if os.path.exists(caminho_temp):
            try:
                df_temp = pd.read_csv(caminho_temp, sep="\t", encoding="utf-16")
                break
            except Exception:
                time.sleep(1)
        else:
            time.sleep(1)

    u.print_log(logs_widget, f"✅ Lote {indice} salvo em {nome_temp}")
    # Limpeza do arquivo
    df_temp = df_temp.drop([0,1,2,4], errors='ignore').reset_index(drop=True)
    df_temp = df_temp.drop(df_temp.columns[0], axis=1)

    # Voltar para preparar próximo lote
    session.findById("wnd[0]").sendVKey(3)

    return df_temp, caminho_temp

# ---------------------------
# Execução completa
# ---------------------------
def executar_refat_massivo(logs_widget, caminho_planilha, periodo, tamanho_lote=2500,
                           pasta_download=PASTA_DOWNLOAD_PADRAO, p_file="/interf", coluna="Instalação"):
    os.makedirs(pasta_download, exist_ok=True)

    # Conectar e abrir transação SAP
    session = conectar_sap()
    abrir_transacao(session, periodo, p_file)

    # Abrir Excel
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel_app.Workbooks.Open(caminho_planilha)
    ws = wb.Sheets(1)
    # Descobrir índice da coluna pelo nome
    cabecalhos = [ws.Cells(1, col).Value for col in range(1, ws.UsedRange.Columns.Count + 1)]
    try:
        col_index = cabecalhos.index(coluna) + 1  # +1 porque Excel é 1-based
    except ValueError:
        wb.Close(False)
        excel_app.Quit()
        raise ValueError(f"Coluna '{coluna}' não encontrada na planilha.")

    # Converter número da coluna em letra
    col_letter = string.ascii_uppercase[col_index - 1]

    # Descobrir última linha preenchida da coluna
    ultima_linha = ws.Cells(ws.Rows.Count, col_letter).End(constants.xlUp).Row

    # Calcular lotes
    total_linhas = ultima_linha - 1  # considerando cabeçalho
    total_lotes = math.ceil(total_linhas / tamanho_lote)

    relatorios = []

    for i in range(total_lotes):
        start_row = 2 + i * tamanho_lote
        end_row = min(1 + (i + 1) * tamanho_lote, ultima_linha)
        df_temp, caminho_temp = processar_lote(
            logs_widget, session, ws, start_row, end_row, i+1,
            pasta_download=pasta_download, coluna=coluna
        )
        relatorios.append((df_temp, caminho_temp))

    wb.Close(False)
    excel_app.Quit()

    return relatorios