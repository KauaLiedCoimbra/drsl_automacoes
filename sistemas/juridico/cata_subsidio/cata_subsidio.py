import os
import utils as u
import pythoncom
import pyperclip
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

def coletar_dados(instalacoes, infos_selecionadas, periodo_inicio, periodo_fim, logs_widget=None, interromper_var=None):
    #==================================
    # INICIALIZAÇÃO
    #==================================
    pythoncom.CoInitialize()
    session = u.conectar_sap()

    inicio = pd.to_datetime(periodo_inicio) 
    fim = pd.to_datetime(periodo_fim)

    u.print_log(logs_widget, f"🔹 Iniciando cata-subsídio para {len(instalacoes)} instalações")
    
    # Para cada instalação...
    for idx, inst in enumerate(instalacoes, 1):
        if interromper_var and interromper_var.get():
            u.print_log(logs_widget, f"⚠️ Interrompido pelo usuário em {inst}")
            break

        info = {}

        u.print_log(logs_widget, f"🔹 [{idx}/{len(instalacoes)}] Processando instalação {inst}")

        # Criação da pasta
        pasta_inst = os.path.join(os.getcwd(), f"instalacao_{inst}")
        os.makedirs(pasta_inst, exist_ok=True)
        caminho_excel = os.path.join(pasta_inst, f"relatorio_{inst}.xlsx")

        #==================================
        # COLETA INFORMAÇÕES GERAIS
        #==================================
        u.abrir_transacao(session, "ES32")

        u.print_log(logs_widget, f"   ↳ Coletando Informações de contrato (ES32)")
        session.findById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = inst
        session.findById("wnd[0]").sendVKey(0)

        # INSTALAÇÃO
        info["instalacao"] = inst
  
        # CONTRATO
        info["contrato"] = session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text
  
        # PN
        info["pn"] = session.findById("wnd[0]/usr/txtEANLD-PARTNER").text
  
        # ENDEREÇO
        info["endereço"] = session.findById("wnd[0]/usr/txtEANLD-LINE1").text
   
        # LOCAL DE CONSUMO
        info["local_consumo"] = session.findById("wnd[0]/usr/ctxtEANLD-VSTELLE").text
    
        # ZONA
        zona_text = session.findById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[6,0]").text
        codigo = zona_text[3:5]
        if codigo == "BU":
            info["zona"] = f"Urbana - ({zona_text})"
        elif codigo == "BR":
            info["zona"] = f"Rural - ({zona_text})"
        elif codigo == "BC":
            info["zona"] = f"Simultânea - ({zona_text})"
        elif codigo == "TR":
            info["zona"] = f"Transitório - ({zona_text})"
        else:
            info["zona"] = f"{zona_text}"
    
        # TARIFA
        info["tarifa"] = f"{session.findById('wnd[0]/usr/ctxtEANLD-ANLART').text} - {session.findById('wnd[0]/usr/txtEANLD-ANLARTTEXT').text}"
      
        # DATA DE LIGAÇÃO
        session.findById("wnd[0]/tbar[1]/btn[34]").press()
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH1/ssubSUB1:SAPLEADS2:0110/cntlCONTROL_AREA1/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH1/ssubSUB1:SAPLEADS2:0110/cntlCONTROL_AREA1/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        data_ligacao = pyperclip.paste()

        df_data_ligacao = u.corrige_na_clipboard(data_ligacao, idx, linhas_remover_i1=[0,1,2,3,5], linhas_remover_padrao=[0,1,2,3,5], colunas_remover=[2, 5, 11])
        df_data_ligacao.columns = df_data_ligacao.iloc[0]
        df_data_ligacao = df_data_ligacao[1:].reset_index(drop=True) 
     
        # CONTA CONTRATO
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/usr/txtEANLD-VERTRAG").setFocus()
        session.findById("wnd[0]").sendVKey(2)
        info["conta_contrato"] = session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLES20:0311/subACCOUNT:SAPLES06:1010/ctxtEVERD-VKONTO").text
      
        # FASE
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select()
        info["fase"] = session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPLES20:0314/txtTE191T-TEXT30").text
      
        # CADASTRO DE E-MAIL
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select()
        df_contatos = pd.DataFrame(columns=["Nome", "Telefone", "Email"])
      
        # CADASTRO DE SMS
      
        # MICRO/MINIGERAÇÃO
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select()
        df_gd = pd.DataFrame(columns=["Campo1", "Campo2"])

        # PIX
        session.findById("wnd[0]/usr/subTABSTRIP:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB08").select()
        df_pix = pd.DataFrame(columns=["PIX"])

        # CONSULTA DE LEITURAS
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[1]/btn[34]").press()
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpSUB9").select()
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpSUB9/ssubSUB9:SAPLEADS2:0190/cntlCONTROL_AREA9/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpSUB9/ssubSUB9:SAPLEADS2:0190/cntlCONTROL_AREA9/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        consulta_leituras = pyperclip.paste()
        df_leituras = u.corrige_na_clipboard(consulta_leituras, idx, linhas_remover_i1=[0, 1, 2, 3, 5], linhas_remover_padrao=[0, 1, 2, 3, 5], colunas_remover=[10, 11, 12])
        df_leituras.columns = df_leituras.iloc[0]
        df_leituras = df_leituras[1:].reset_index(drop=True) 
        df_leituras["Dt.leitura"] = pd.to_datetime(df_leituras["Dt.leitura"], format="%d.%m.%Y", errors="coerce")
        df_leituras_filtradas = df_leituras[(df_leituras["Dt.leitura"] >= inicio) & (df_leituras["Dt.leitura"] <= fim)].reset_index(drop=True)

        # FATURAS
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH2").select
        
        shell = session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH2/"
                                "ssubSUB2:SAPLEADS2:0120/cntlCONTROL_AREA2/shellcont/shell")

        coluna_opbel = [shell.getCellValue(row, "OPBEL") for row in range(shell.RowCount)]
        coluna_datas = [shell.getCellValue(row, "BEGABRPE") for row in range(shell.RowCount)]
        
        for i in range(1, len(coluna_datas)):
            if coluna_datas[i].strip() == "":
                coluna_datas[i] = coluna_datas[i-1]

        coluna_datas_dt = [datetime.strptime(d, "%d.%m.%Y") for d in coluna_datas]

        valores_filtrados = [
            opbel for opbel, data in zip(coluna_opbel, coluna_datas_dt)
            if inicio <= data <= fim
        ]
        faturas_clipboard = "\n".join(valores_filtrados)
        pyperclip.copy(faturas_clipboard)

        u.abrir_transacao(session, "ZCCSFAT104")

        session.findById("wnd[0]/usr/ctxtSC_DCIMP-LOW").text = "1"
        session.findById("wnd[0]/usr/btn%_SC_DCIMP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[24]").press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        #==================================
        # RELATÓRIO EXCEL
        #==================================
        with pd.ExcelWriter(caminho_excel, engine='xlsxwriter') as writer:
            # Aba Informações Gerais
            df_geral = pd.DataFrame([{
                "Instalação": info["instalacao"],
                "Contrato atual": info["contrato"],
                "PN": info["pn"],
                "Conta Contrato": info["conta_contrato"],
                "Fase": info["fase"],
                "Endereço": info["endereço"],
                "Local de consumo": info["local_consumo"],
                "Zona": info["zona"],
                "Tarifa": info["tarifa"]
            }])
            df_geral.to_excel(writer, sheet_name='Informações Gerais', index=False)

            # Aba Data de ligação
            if not df_data_ligacao.empty:
                df_data_ligacao.to_excel(writer, sheet_name='Data de ligação', index=False)

            # Aba Leituras
            if not df_leituras_filtradas.empty:
                df_leituras_filtradas.to_excel(writer, sheet_name='Leituras', index=False)

            # Aba Contatos
            df_contatos.to_excel(writer, sheet_name='Contatos', index=False)

            # Aba GD
            df_gd.to_excel(writer, sheet_name='GD', index=False)

            # Aba PIX
            df_pix.to_excel(writer, sheet_name='PIX', index=False)

        wb = load_workbook(caminho_excel)
        for sheet in wb.worksheets:
            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter  # pega a letra da coluna
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = max_length + 2  # +2 para dar uma folga
                sheet.column_dimensions[column].width = adjusted_width
        wb.save(caminho_excel)

    u.print_log(logs_widget, "✅ Coleta de dados finalizada!")
