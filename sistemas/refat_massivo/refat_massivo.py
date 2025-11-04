import os
import time
import pandas as pd
import pyperclip
import win32com.client as win32
import utils as u
import string
from win32com.client import constants

PASTA_DOWNLOAD_PADRAO = r"C:\Users\2038860\OneDrive - CPFL Energia S A\projetos\automatron\sistemas\refat_massivo\relatorios"

# ---------------------------
# Abrir transação
# ---------------------------
def configura_refat(session, periodo, p_file):
    session.findById("wnd[0]/usr/radP_ACES2").select()
    session.findById("wnd[0]/usr/txtS_BPER-LOW").text = periodo
    session.findById("wnd[0]/usr/ctxtP_FILE").text = p_file

# ---------------------------
# Ler coluna de interesse
# ---------------------------
def ler_coluna_excel(logs_widget, caminho_planilha, coluna_nome):
    # Abrir Excel invisível
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    excel_app.Visible = False
    wb = excel_app.Workbooks.Open(caminho_planilha)
    ws = wb.Sheets(1)

    # Descobrir índice da coluna pelo nome
    cabecalhos = [ws.Cells(1, col).Value for col in range(1, ws.UsedRange.Columns.Count + 1)]
    try:
        col_index = cabecalhos.index(coluna_nome) + 1  # +1 porque Excel é 1-based
    except ValueError:
        wb.Close(False)
        excel_app.Quit()
        raise ValueError(f"Coluna '{coluna_nome}' não encontrada na planilha.")

    # Converter número da coluna em letra
    col_letter = string.ascii_uppercase[col_index - 1]

    # Descobrir última linha preenchida da coluna
    ultima_linha = ws.Cells(ws.Rows.Count, col_letter).End(constants.xlUp).Row

    # Extrair valores da coluna de uma vez (lista de strings)
    range_valores = ws.Range(f"{col_letter}2:{col_letter}{ultima_linha}")
    valores_raw = range_valores.Value

    # Normalizar para lista de listas
    if not isinstance(valores_raw, tuple):
        valores_raw = ((valores_raw,),)

    valores = [str(c[0]) for c in valores_raw if c[0] is not None]

    u.print_log(logs_widget, f"✔️ Coluna '{coluna_nome}' carregada ({len(valores)} valores).")

    # Retorna ws, letra da coluna e lista de valores
    return ws, col_letter, valores, excel_app, wb

# ---------------------------
# Processar lotes com cópia direta do Excel (com flag de interrupção)
# ---------------------------
def processar_lotes(logs_widget, session, ws, col_letter, valores, tamanho_lote, interromper_flag=None):
    lotes = [valores[i:i + tamanho_lote] for i in range(0, len(valores), tamanho_lote)]
    df_final = pd.DataFrame()

    for i, lote in enumerate(lotes, 1):
        # Verifica interrupção
        if interromper_flag and interromper_flag.get():
            u.print_log(logs_widget, "⏹ Execução interrompida pelo usuário. Salvando progresso parcial...")
            break

        u.print_log(logs_widget, f"🔹 Processando lote {i}/{len(lotes)} ({len(lote)} itens)...")

        start_row = 2 + (i - 1) * tamanho_lote
        end_row = start_row + len(lote) - 1

        # Abrir múltipla seleção SAP
        session.findById("wnd[0]/usr/btn%_S_ANLG_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()  # limpar antigo

        # Copiar intervalo do Excel
        intervalo = ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}")
        intervalo.Copy()

        # Colar no SAP
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  # confirmar

        # Executar consulta
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Exportar via clipboard
        session.findById("wnd[0]").sendVKey(45)
        session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)

        # Ler dados do clipboard
        texto = pyperclip.paste()
        df_final, df_lote = u.corrige_na_clipboard(texto, i)
        
        u.print_log(logs_widget, f"✅ Lote {i} concluído ({len(df_lote)} linhas úteis).")

        # Voltar SAP
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)

    return df_final

# ---------------------------
# Função principal
# ---------------------------
def executar_refat_massivo(logs_widget, caminho_planilha, periodo, tamanho_lote,
                           p_file, coluna, interromper_flag=None, pasta_download=PASTA_DOWNLOAD_PADRAO):
    try:
        u.print_log(logs_widget, "🔗 Conectando ao SAP...")
        session = u.conectar_sap()

        u.print_log(logs_widget, f"🧭 Acessando transação com período {periodo} e arquivo {p_file}...")
        u.abrir_transacao(session, "ZFAT0657")
        configura_refat(session, periodo, p_file)

        u.print_log(logs_widget, "📖 Lendo planilha...")

        # ✅ Leitura Excel otimizada
        ws, col_letter, valores, excel_app, wb = ler_coluna_excel(logs_widget, caminho_planilha, coluna)

        u.print_log(logs_widget, f"📦 Total de {len(valores)} registros encontrados na coluna '{coluna}'.")
        u.print_log(logs_widget, f"⚙️ Iniciando processamento em lotes de {tamanho_lote}...")

        # Processar lotes diretamente do Excel, com flag de interrupção
        df_final = processar_lotes(logs_widget, session, ws, col_letter, valores, tamanho_lote, interromper_flag=interromper_flag)

        # Fechar Excel ao final
        wb.Close(False)
        excel_app.Quit()

        # Gerar caminho final automático
        nome_final = os.path.splitext(os.path.basename(caminho_planilha))[0] + "_resultado.xlsx"
        caminho_final = os.path.join(pasta_download, nome_final)

        if df_final.empty:
            u.print_log(logs_widget, "⚠️ Nenhum dado processado. Nenhum arquivo será salvo.")
            return None

        u.print_log(logs_widget, "💾 Salvando resultado final...")

        for tentativa in range(3):
            try:
                df_final.to_excel(caminho_final, index=False)
                u.print_log(logs_widget, f"🎉 Processamento concluído!\n📂 Arquivo salvo em:\n{caminho_final}")
                break
            except PermissionError:
                if tentativa < 2:
                    u.print_log(logs_widget, f"⚠️ Arquivo em uso ({tentativa+1}/3). Tentando novamente...")
                    time.sleep(2)
                else:
                    base, ext = os.path.splitext(caminho_final)
                    caminho_final = f"{base}_{int(time.time())}{ext}"
                    df_final.to_excel(caminho_final, index=False)
                    u.print_log(logs_widget, f"⚠️ Salvamento alternativo:\n{caminho_final}")
                    break

        return caminho_final

    except Exception as e:
        u.print_log(logs_widget, f"❌ Erro durante execução: {e}")
        raise

