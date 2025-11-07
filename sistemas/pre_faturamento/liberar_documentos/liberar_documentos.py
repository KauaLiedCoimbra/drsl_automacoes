import time
import utils as u
import pythoncom

def executar_liberar_documentos(logs_widget, matricula="6328490", layout="//CHARLES"):
    pythoncom.CoInitialize()
    try:
        u.print_log(logs_widget, "🔗 Conectando ao SAP...")
        session = u.conectar_sap()
        if not session:
            raise Exception("Não foi possível conectar ao SAP.")

        u.print_log(logs_widget, "🧭 Abrindo transação EA05...")
        u.abrir_transacao(session, "EA05")
        session.findById("wnd[0]").sendVKey(17)
        session.findById("wnd[1]/usr/txtENAME-LOW").text = matricula
        session.findById("wnd[0]").sendVKey(8)

        # Selecionar variante
        grid_popup = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")

        while True:
            try:
                if grid_popup.RowCount > 0:
                    break
            except:
                pass
            time.sleep(0.2)

        # SELECIONAR VARIANTE
        grid_popup.currentCellRow = 1
        grid_popup.selectedRows = "1"
        grid_popup.doubleClickCurrentCell

        session.findById("wnd[1]").sendVKey (2)
        session.findById("wnd[0]").sendVKey (8)

        # SELECIONAR LAYOUT
        session.findById("wnd[0]").sendVKey (33)
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem("&FILTER")
        session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "//CHARLES"
        session.findById("wnd[2]").sendVKey (0)

        while True:
            try:
                layout = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
                if layout.RowCount >= 0:
                    break
            except:
                time.sleep(0.2)

        layout.currentCellRow = 0     
        layout.selectedRows = "0"    
        layout.doubleClickCurrentCell()

        # FILTRAR AMOUNT1
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("VALIDATION")
        session.findById("wnd[0]").sendVKey (29)
        session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "AMOUNT1"
        session.findById("wnd[1]").sendVKey (0)

        #SELECIONA AS LINHAS COM DESVIO ENTRE 0 E 1000
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        grid.selectColumn("DEVIATION")

        num_linhas = grid.RowCount
        linhas_para_selecionar = []

        while True:
            first = grid.firstVisibleRow
            last = first + grid.visibleRowCount
            if last > num_linhas:
                last = num_linhas

            for i in range(first, last):
                try:
                    valor = grid.GetCellValue(i, "DEVIATION")
                    if valor and valor.strip() != "":
                        valor_num = float(valor.replace(",", "."))
                        if 0 < valor_num < 1000:
                            linhas_para_selecionar.append(i)
                except:
                    continue

            if last >= num_linhas:
                break

            grid.firstVisibleRow = last

        if linhas_para_selecionar:
            grid.selectedRows = ",".join(str(x) for x in linhas_para_selecionar)

        # Liberar documentos
        session.findById("wnd[0]").sendVKey(16)
        for _ in range(len(linhas_para_selecionar)):
            session.findById("wnd[1]").sendVKey(0)

        u.print_log(
            logs_widget,
            f"✅ Documentos liberados: {len(linhas_para_selecionar)} linhas selecionadas"
        )

    except Exception as e:
        u.print_log(logs_widget, f"❌ Erro: {e}")
        raise
