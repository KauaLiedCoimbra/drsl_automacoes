import win32com.client

def transcrever_sap_linear(print_log):
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not SapGuiAuto:
        print_log("SAP GUI não está rodando")
        return
    
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    linhas = []

    def percorrer_objetos(obj, caminho="app"):
        try:
            tipo_obj = type(obj).__name__
            obj_id = getattr(obj, 'Id', '')
            obj_name = getattr(obj, 'Name', '')
            obj_text = getattr(obj, 'Text', '')

            # Monta a linha no formato desejado
            caminho_completo = f"{caminho} ({tipo_obj})"
            if obj_name:
                caminho_completo += f" -> {obj_name}"
            if obj_text:
                caminho_completo += f" -> {obj_text}"

            linhas.append(caminho_completo)

            # Para grids/tabelas, percorre células (opcional)
            if tipo_obj in ['GuiGridView', 'GuiTableControl']:
                for r in range(obj.RowCount):
                    for c in range(obj.ColumnCount):
                        try:
                            val = obj.GetCellValue(r, c)
                        except:
                            val = ""
                        linhas.append(f"{caminho}/{obj_name}[{r},{c}] (Cell) -> {val}")

            # Percorre filhos recursivamente
            for i in range(obj.Children.Count):
                percorrer_objetos(obj.Children(i), caminho + f"/{obj_name}[{i}]")

        except Exception:
            pass

    percorrer_objetos(session, caminho="/app/con[0]/ses[0]/wnd[0]")

    # Salva em TXT
    with open("sap_transcricao_linear.txt", "w", encoding="utf-8") as f:
        for linha in linhas:
            f.write(linha + "\n")
    
    print_log(f"Transcrição linearizada concluída. {len(linhas)} linhas escritas.")