import time
import os
import utils as u

def coletar_dados(instalacoes, infos_selecionadas, periodo_inicio, periodo_fim, logs_widget=None, interromper_var=None):
    """
    Para cada instalação:
        - Acessa as transações SAP correspondentes
        - Coleta as informações desejadas
        - Salva em arquivos/pastas estruturadas
    """
    session = u.conectar_sap()

    u.print_log(logs_widget, f"🔹 Iniciando coleta de dados para {len(instalacoes)} instalações")

    for idx, inst in enumerate(instalacoes, 1):
        if interromper_var and interromper_var.get():
            u.print_log(logs_widget, f"⚠️ Interrompido pelo usuário em {inst}")
            break

        u.print_log(logs_widget, f"🔹 [{idx}/{len(instalacoes)}] Processando instalação {inst}")

        # Exemplo de estrutura: criar pasta da instalação
        pasta_inst = os.path.join(os.getcwd(), f"instalacao_{inst}")
        os.makedirs(pasta_inst, exist_ok=True)

        # Aqui você chamaria funções específicas para cada transação SAP
        if "Histórico de consumo" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "ZCCSPEC015"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar ZCCSPEC015 e salvar resultados
            u.print_log(logs_widget, f"   ↳ Coletando Histórico de consumo (ZCCSPEC015)")

        if "Faturas / Pagamentos / Parcelamento" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "FPL9"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar FPL9 e salvar resultados
            u.print_log(logs_widget, f"   ↳ Coletando Faturas / Pagamentos / Parcelamento (FPL9)")

        if "Devoluções de créditos" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "FPL9"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar FPL9 para devoluções
            u.print_log(logs_widget, f"   ↳ Coletando Devoluções de créditos (FPL9)")

        if "Negativações / Protesto" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "ZCCSACC064"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar ZCCSACC064
            u.print_log(logs_widget, f"   ↳ Coletando Negativações / Protesto (ZCCSACC064)")

        if "Faturas em PDF" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "ZCCSFAT104"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar ZCCSFAT104 e salvar PDFs
            u.print_log(logs_widget, f"   ↳ Coletando Faturas em PDF (ZCCSFAT104)")

        if "Consulta de leituras" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "ES32"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar ES32
            u.print_log(logs_widget, f"   ↳ Coletando Consulta de leituras (ES32)")

        if "Informações de contrato" in infos_selecionadas:
            session.findById("wnd[0]/tbar[0]/okcd").text = "ES32"
            session.findById("wnd[0]").sendVKey(0)
            # TODO: acessar ES32 novamente para dados de contrato
            u.print_log(logs_widget, f"   ↳ Coletando Informações de contrato (ES32)")
        # Simular tempo de processamento
        time.sleep(0.5)

    u.print_log(logs_widget, "✅ Coleta de dados finalizada!")
