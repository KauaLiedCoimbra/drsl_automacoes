import win32com.client


def transcrever_sap_linear(print_log, arquivo_saida="sap_tela_detalhada.txt"):
    try:
        print_log("▶ Iniciando transcrição da tela SAP...")

        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            print_log("❌ Não foi possível acessar o SAP GUI. Verifique se o SAP está aberto e com scripting habilitado (Alt+F12 → Opções → Scripting).")
            return

        application = SapGuiAuto.GetScriptingEngine
        if application is None or application.Children.Count == 0:
            print_log("❌ Nenhuma conexão SAP ativa encontrada. Abra o SAP e entre em uma sessão antes de executar.")
            return

        connection = application.Children(0)
        if connection.Children.Count == 0:
            print_log("❌ Nenhuma sessão aberta no SAP. Entre em um sistema (ex: SE16) e tente novamente.")
            return

        session = connection.Children(0)
        window = session.ActiveWindow
        print_log("✅ Conectado ao SAP GUI com sucesso.")


        def percorrer_elementos(elemento, nivel=0):
            linhas = []
            try:
                tipo = getattr(elemento, "Type", "Desconhecido")
                texto = getattr(elemento, "Text", "")
                linhas.append("  " * nivel + f"{elemento.Id} ({tipo}) -> {texto}")

                # Captura conteúdo de tabelas
                if tipo == "GuiGridView":
                    colunas = elemento.ColumnCount
                    linhas.append("  " * (nivel + 1) + f"--- Conteúdo da Tabela ({colunas} colunas) ---")
                    for row in range(elemento.RowCount):
                        celulas = [str(elemento.GetCellValue(row, col)) for col in range(colunas)]
                        linhas.append("  " * (nivel + 1) + f"Linha {row}: " + " | ".join(celulas))
            except Exception:
                pass

            children = getattr(elemento, "Children", None)
            if children:
                for child in children:
                    linhas.extend(percorrer_elementos(child, nivel + 1))
            return linhas

        conteudo = percorrer_elementos(window)

        with open(arquivo_saida, "w", encoding="utf-8") as f:
            f.write("\n".join(conteudo))

        print_log(f"💾 Transcrição salva em: {arquivo_saida}")
        print_log("✅ Processo concluído com sucesso.")
    except Exception as e:
        print_log(f"❌ Erro durante execução: {e}")
