import time
import os
import utils as u
import pythoncom
import pyperclip

def coletar_dados(instalacoes, infos_selecionadas, periodo_inicio, periodo_fim, logs_widget=None, interromper_var=None):
    """
    Para cada instalação:
        - Acessa as transações SAP correspondentes
        - Coleta as informações desejadas
        - Salva em arquivos/pastas estruturadas
    """
    pythoncom.CoInitialize()
    session = u.conectar_sap()

    u.print_log(logs_widget, f"🔹 Iniciando coleta de dados para {len(instalacoes)} instalações")

    for idx, inst in enumerate(instalacoes, 1):
        if interromper_var and interromper_var.get():
            u.print_log(logs_widget, f"⚠️ Interrompido pelo usuário em {inst}")
            break

        info = {}

        u.print_log(logs_widget, f"🔹 [{idx}/{len(instalacoes)}] Processando instalação {inst}")

        # Exemplo de estrutura: criar pasta da instalação
        pasta_inst = os.path.join(os.getcwd(), f"instalacao_{inst}")
        os.makedirs(pasta_inst, exist_ok=True)
        print(inst)
        # Aqui você chamaria funções específicas para cada transação SAP
        #==================================
        # DADOS BÁSICOS
        #==================================
        u.abrir_transacao(session, "ES32")

        u.print_log(logs_widget, f"   ↳ Coletando Informações de contrato (ES32)")
        session.findById("wnd[0]/usr/ctxtEANLD-ANLAGE").text = inst
        session.findById("wnd[0]").sendVKey(0)

        # CONTRATO
        info = {"contrato": session.findById("wnd[0]/usr/txtEANLD-VERTRAG").text}
        print(info["contrato"])

        # PN
        info = {"pn": session.findById("wnd[0]/usr/txtEANLD-PARTNER").text}
        print(info["pn"])

        # ENDEREÇO
        info = {"endereço": session.findById("wnd[0]/usr/txtEANLD-LINE1").text}
        print(info["endereço"])

        # LOCAL DE CONSUMO
        info = {"local_consumo": session.findById("wnd[0]/usr/ctxtEANLD-VSTELLE").text}
        print(info["local_consumo"])

        # ZONA
        zona_text = session.findById("wnd[0]/usr/tblSAPLES30TC_TIMESL/ctxtEANLD-ABLEINH[6,0]").text
        codigo = zona_text[3:5]
        if codigo == "BU":
            info = {"zona": f"Urbana ({zona_text})"}
        elif codigo == "BR":
            info = {"zona": f"Rural ({zona_text})"}
        elif codigo == "BC":
            info = {"zona": f"Simultânea ({zona_text})"}
        elif codigo == "TR":
            info = {"zona": f"Transitório ({zona_text})"}
        else:
            info = {"zona": zona_text}
        print(info["zona"])

        # FASE
        tp_instal = session.findById("wnd[0]/usr/ctxtEANLD-ANLART").text
        if tp_instal == "0001":
            info = {"fase": f"Monofásico"}
        if tp_instal == "0002":
            info = {"fase": f"Bifásico"}
        if tp_instal == "0003":
            info = {"fase": f"Trifásico"}
        print(info["fase"])

        # DATA DE LIGAÇÃO
        session.findById("wnd[0]/tbar[1]/btn[34]").press()
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH1/ssubSUB1:SAPLEADS2:0110/cntlCONTROL_AREA1/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH1/ssubSUB1:SAPLEADS2:0110/cntlCONTROL_AREA1/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]/usr/sub:SAPLSPO5:0201/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        data_ligacao = pyperclip.paste()

        # FATURAMENTO
        session.findById("wnd[0]/usr/tabsMYTABSTRIP/tabpPUSH2").select()
        time.sleep(2)

        if "Informações de contrato" in infos_selecionadas:
            u.abrir_transacao(session, "ES32")
            time.sleep(2)
            # TODO: acessar ES32 e salvar resultados
        if "Histórico de consumo" in infos_selecionadas:
            u.abrir_transacao(session, "ZCCSPEC015")
            time.sleep(2)
            # TODO: acessar ZCCSPEC015 e salvar resultados
            u.print_log(logs_widget, f"   ↳ Coletando Histórico de consumo (ZCCSPEC015)")

        if "Faturas / Pagamentos / Parcelamento" in infos_selecionadas:
            u.abrir_transacao(session, "FPL9")
            time.sleep(2)
            # TODO: acessar FPL9 e salvar resultados
            u.print_log(logs_widget, f"   ↳ Coletando Faturas / Pagamentos / Parcelamento (FPL9)")

        if "Devoluções de créditos" in infos_selecionadas:
            u.abrir_transacao(session, "FPL9")
            time.sleep(2)
            # TODO: acessar FPL9 para devoluções
            u.print_log(logs_widget, f"   ↳ Coletando Devoluções de créditos (FPL9)")

        if "Negativação" in infos_selecionadas:
            u.abrir_transacao(session, "ES16N")
            time.sleep(2)
            # TODO: acessar FPL9 para devoluções
            u.print_log(logs_widget, f"   ↳ Coletando Devoluções de créditos (FPL9)")

        if "Negativações / Protesto" in infos_selecionadas:
            u.abrir_transacao(session, "ZCCSACC064")
            time.sleep(2)
            # TODO: acessar ZCCSACC064
            u.print_log(logs_widget, f"   ↳ Coletando Negativações / Protesto (ZCCSACC064)")

        if "Faturas em PDF" in infos_selecionadas:
            u.abrir_transacao(session, "ZCCSFAT104")
            time.sleep(2)
            # TODO: acessar ZCCSFAT104 e salvar PDFs
            u.print_log(logs_widget, f"   ↳ Coletando Faturas em PDF (ZCCSFAT104)")

        if "Consulta de leituras" in infos_selecionadas:
            u.abrir_transacao(session, "ES32")
            time.sleep(2)
            # TODO: acessar ES32
            u.print_log(logs_widget, f"   ↳ Coletando Consulta de leituras (ES32)")
        # Simular tempo de processamento
        time.sleep(0.5)

    u.print_log(logs_widget, "✅ Coleta de dados finalizada!")
