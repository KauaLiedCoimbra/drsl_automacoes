import os
import time
import math
import pandas as pd
import win32com.client as win32
from win32com.client import constants

# === CONFIGURAÇÕES ===
CAMINHO_PLANILHA = r"C:\Users\2038860\OneDrive - CPFL Energia S A\projetos\refat massivo\baixarenda.xlsx"
COLUNA_INSTALACOES = "Instalação"
TAMANHO_LOTE = 2500
PASTA_DOWNLOAD = r"C:\Users\2038860\OneDrive - CPFL Energia S A\projetos\refat massivo\relatorios_temp"
ARQUIVO_FINAL = r"C:\Users\2038860\OneDrive - CPFL Energia S A\projetos\refat massivo"
PERIODO = "2025/10"
P_FILE = "/interf"

os.makedirs(PASTA_DOWNLOAD, exist_ok=True)

# === CONECTAR AO SAP ===
def conectar_sap():
    SapGuiAuto = win32.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    return session

# === ABRIR TRANSAÇÃO ZFAT0657 UMA VEZ ===
def abrir_transacao(session):
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFAT0657"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radP_ACES2").select()
    session.findById("wnd[0]/usr/txtS_BPER-LOW").text = PERIODO
    session.findById("wnd[0]/usr/ctxtP_FILE").text = P_FILE
    session.findById("wnd[0]/usr/ctxtP_FILE").setFocus()
    session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition = len(P_FILE)

# === PROCESSAR UM LOTE ===
def processar_lote(session, ws, start_row, end_row, indice):
    # Abrir múltipla seleção
    session.findById("wnd[0]/usr/btn%_S_ANLG_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()  # Limpar conteúdo antigo

    # Copiar intervalo do Excel
    intervalo = ws.Range(f"C{start_row}:C{end_row}")
    intervalo.Copy()

    # Colar no SAP
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Confirmar seleção

    # Executar consulta
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Exportar direto para Excel sem abrir tela
    nome_temp = f"relatorio_temp_{indice}.XLS"
    session.findById("wnd[0]").sendVKey(45)
    session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = PASTA_DOWNLOAD
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_temp
    session.findById("wnd[0]").sendVKey(0)
    print('Arquivo baixado')
    time.sleep(2)

    # Esperar arquivo ser criado
    caminho_temp = os.path.join(PASTA_DOWNLOAD, nome_temp)
    while True:
        if os.path.exists(caminho_temp):
            try:
                # Forçar engine compatível com XLS antigo
                df_temp = pd.read_csv(caminho_temp, sep="\t", encoding="utf-16")
                break
            except Exception as e:
                print(f"Arquivo encontrado, mas ainda não pronto: {e}")
                time.sleep(1)
        else:
            print("Arquivo ainda não existe")
            time.sleep(1)

    print(f"✅ Lote {indice} salvo em {nome_temp}")

    # Apagar linhas 1,2,3,5
    df_temp = df_temp.drop([0,1,2,4], errors='ignore').reset_index(drop=True)

    # Apagar coluna A
    df_temp = df_temp.drop(df_temp.columns[0], axis=1)

    # Voltar para preparar próximo lote
    session.findById("wnd[0]").sendVKey(3)

    return df_temp

# === EXECUÇÃO PRINCIPAL ===
if __name__ == "__main__":
    # Conectar SAP e abrir transação
    session = conectar_sap()
    abrir_transacao(session)

    # Abrir Excel
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel_app.Workbooks.Open(CAMINHO_PLANILHA)
    ws = wb.Sheets(1)
    ultima_linha = ws.Cells(ws.Rows.Count, "C").End(constants.xlUp).Row

    # Calcular lotes
    total_linhas = ultima_linha - 1  # considerando A1 como cabeçalho
    total_lotes = math.ceil(total_linhas / TAMANHO_LOTE)
    print(f"Total de lotes: {total_lotes}")

    for i in range(total_lotes):
        start_row = 2 + i * TAMANHO_LOTE
        end_row = min(1 + (i + 1) * TAMANHO_LOTE, ultima_linha)
        print(f"🔹 Processando lote {i+1}/{total_lotes} (linhas {start_row}-{end_row})...")
        try:
            df_temp = processar_lote(session, ws, start_row, end_row, i+1)
        except Exception as e:
            print(f"⚠️ Erro no lote {i+1}: {e}")

    # Fechar Excel
    wb.Close(False)
    excel_app.Quit()