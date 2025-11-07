import win32com.client as win32
import pyperclip
import time
import pandas as pd
from datetime import datetime
import utils as u
import pythoncom

def executar_notas_diarias(destinatario: str, logs_widget=None, interromper_var=None):
    
    pythoncom.CoInitialize()

    def log(msg):
        if logs_widget:
            u.print_log(logs_widget, msg)
        else:
            print(msg)

    if not destinatario.strip():
        log("❌ Destinatário não informado.")
        return

    if interromper_var and interromper_var.get():
        log("⚠️ Execução interrompida antes do início.")
        return

    log("🔹 Conectando ao SAP...")
    try:
        SapGuiAuto = win32.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        log(f"❌ Erro ao conectar ao SAP: {e}")
        return

    # ===============================
    # Coleta de informações do SAP
    # ===============================
    try:
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "0000000003"
        session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("0000000003")
        session.findById("wnd[1]/usr/lbl[22,4]").caretPosition = 4
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[0]/usr/subAREA04:SAPLCRM_CIC_SLIM_ACTION_BOX:0110/cntlCRMCICSLIMABCONTAINER/shellcont/shell").pressContextButton("ZSER")
        session.findById("wnd[0]/usr/subAREA04:SAPLCRM_CIC_SLIM_ACTION_BOX:0110/cntlCRMCICSLIMABCONTAINER/shellcont/shell").selectContextMenuItem("FPL9")
        while connection.Children.Count < 2:
            time.sleep(0.2)
        session = connection.Children(1)
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW58"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        session.findById("wnd[1]/usr/txtV-LOW").text = "CT UNIFICADO"
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").select()
        session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        log("✅ Dados coletados do SAP.")
    except Exception as e:
        log(f"❌ Erro ao coletar dados do SAP: {e}")
        return

    # ===============================
    # Processamento da planilha
    # ===============================
    try:
        tabela = pyperclip.paste()
        linhas = [linha.strip("|").split("|") for linha in tabela.splitlines() if linha.strip()]
        df = pd.DataFrame(linhas)
        primeira = df.iloc[0].astype(str).str.replace(r"[ ,\.]", "", regex=True)
        if all(c.isdigit() for c in primeira):
            df = df.iloc[1:].reset_index(drop=True)
        linhas_para_remover = [0,1,2,4] + [df.index[-1]]
        df = df.drop(linhas_para_remover, errors="ignore").reset_index(drop=True)
        df_limpo = df.reset_index(drop=True)
        df_limpo.columns = df_limpo.iloc[0]
        df_limpo = df_limpo[1:].reset_index(drop=True)

        df_sel = df_limpo.iloc[:, [0, 3]].copy()
        df_sel.columns = ['Concl.desj', 'Nota']
        for col in df_sel.columns:
            df_sel[col] = df_sel[col].astype(str).str.strip()

        df_sel['Concl.desj'] = pd.to_datetime(df_sel['Concl.desj'], dayfirst=True, errors='coerce')
        hoje = pd.Timestamp(datetime.today().date())

        df_ativas = df_sel[df_sel['Concl.desj'] >= hoje].copy()
        df_antigas = df_sel[df_sel['Concl.desj'] < hoje].copy()
        df_ativas['Concl.desj'] = df_ativas['Concl.desj'].dt.strftime('%d.%m.%Y')
        df_antigas['Concl.desj'] = df_antigas['Concl.desj'].dt.strftime('%d.%m.%Y')

        # Tabela Ativas
        tabela_ativas = df_ativas.groupby('Concl.desj')['Nota'].count().to_frame().T
        tabela_ativas['Total Geral'] = tabela_ativas.sum(axis=1)
        tabela_ativas.index = ['Qtd. Notas']
        # linha extra vazia
        linha_vazia = pd.DataFrame([['Roberta' if i!=0 else '' for i in tabela_ativas.columns]], columns=tabela_ativas.columns, index=[''])
        tabela_ativas = pd.concat([tabela_ativas, linha_vazia])

        # Tabela Antigas
        tabela_antigas = df_antigas.groupby('Concl.desj')['Nota'].count().reset_index()
        tabela_antigas.columns = ['Data', 'Notas internas']
        total_geral = pd.DataFrame({'Data': ['Total Geral'], 'Notas internas': [tabela_antigas['Notas internas'].sum()]})
        tabela_antigas = pd.concat([tabela_antigas, total_geral], ignore_index=True)

        log("✅ Planilhas processadas.")
    except Exception as e:
        log(f"❌ Erro ao processar planilha: {e}")
        return

    # ===============================
    # Envio pelo Outlook
    # ===============================
    try:
        html_ativas = tabela_ativas.to_html(index=True, border=1)
        html_antigas = tabela_antigas.to_html(index=False, border=1)

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = "Demandas CT"
        mail.HTMLBody = f"""
        <p>Bom dia!</p>
        <p>Seguem as demandas do CT atualizadas em {hoje.strftime('%d.%m.%Y')}.</p>
        <h3>Notas Ativas</h3>
        {html_ativas}
        <h3>Notas Internas</h3>
        {html_antigas}
        <p>Atenciosamente,</p>
        """
        mail.Send()
        log(f"✅ E-mail enviado para {destinatario}")
    except Exception as e:
        log(f"❌ Erro ao enviar e-mail: {e}")
