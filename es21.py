import win32com.client
import pandas as pd
import utils as u

interrompido = False

def executar_es21(caminho_planilha, print_log):
    """
    Executa a automação ES21.
    - caminho_planilha: caminho completo da planilha de contratos.
    - print_log: função para exibir logs na interface (substitui print()).
    """
    global interrompido
    todos_registros = []

    def salvar_colheita(df_colheita, todos_registros, print_log):
        if not todos_registros:
            return
        df_colheita_save = pd.concat([df_colheita, pd.DataFrame(todos_registros)], ignore_index=True)
        try:
            with pd.ExcelWriter(f"dados_coletados.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_colheita_save.to_excel(writer, sheet_name="Coleta", index=False)
            print_log("✅ Dados salvos em 'dados_coletados.xlsx'")
        except FileNotFoundError:
            df_colheita_save.to_excel("dados_coletados.xlsx", index=False)
            print_log("✅ Arquivo 'dados_coletados.xlsx' criado do zero.")

    # Lê planilhas
    try:
        df = pd.read_excel(caminho_planilha)
    except Exception as e:
        print_log(f"❌ Erro ao abrir a planilha {caminho_planilha}: {e}")
        return
    
    try:
        df_colheita = pd.read_excel("dados_coletados.xlsx")
    except FileNotFoundError:
        df_colheita = pd.DataFrame(columns=['Instalacao','Contrato','RE','Data','VAL.ANTIGO:','VAL.NOVO:'])

    # Corrige colunas
    for col in ['INSTALACAO','CONTRATOS', 'MOTIVO']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: str(int(x)) if isinstance(x, float) else str(x)).str.strip()
        else:
            print_log(f"Coluna {col} não encontrada na planilha.")
            return

    # Conexão SAP
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print_log(f"❌ Erro ao conectar ao SAP: {e}")
        return

    # Maximiza e acessa ES21
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "es21"
    session.findById("wnd[0]").sendVKey(0)
    
    total_contratos = len(df)

    try:
        for index, row in df.iterrows():
            instalacao = row['INSTALACAO']
            contrato = row['CONTRATOS']
            motivo = row["MOTIVO"].zfill(2)
            contratos_restantes = total_contratos - (index + 1)
            print_log(f'🔍 Processando contrato {contrato}... Motivo: {motivo}')

            print(interrompido)
            if interrompido:
                print_log(f"⚠ Execução interrompida pelo usuário. Salvando dados coletados até agora...")
                break

            # Pesquisa contrato
            session.findById("wnd[0]/usr/ctxtEVERD-VERTRAG").text = contrato
            session.findById("wnd[0]/usr/ctxtEVERD-VERTRAG").caretPosition = len(contrato)
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(19)
            session.findById("wnd[0]").sendVKey(47)

            # Inicializa variáveis
            motivo_encontrado = False
            registros = []
            data_atual = None
            re_atual = None

            while not motivo_encontrado:
                if interrompido:
                    break

                # Mapeia elementos
                usr = session.findById("wnd[0]/usr")
                todos = []
                for i in range(usr.Children.Count):
                    child = usr.Children.Item(i)
                    try:
                        texto = child.Text if hasattr(child, "Text") else ""
                        if texto:
                            todos.append({"texto": texto, "top": child.Top, "left": child.Left})
                    except Exception:
                        continue

                for i, elem in enumerate(todos):
                    if interrompido:
                        break
                    texto = elem["texto"]

                    # Quando encontrar uma nova data, atualiza o RE e a data em uso
                    if u.is_data(texto) and not (i > 0 and todos[i - 1]["texto"] in ["Val.antigo:", "Val.novo:"]):
                        if texto != data_atual:
                            data_atual = texto
                            re_atual = todos[i - 1]["texto"] if i > 0 else ""
                        
                    # captura pares de valores dentro da mesma data
                    if texto == "Val.antigo:" and i + 1 < len(todos):
                        val_antigo = todos[i + 1]["texto"]
                        if val_antigo == "Val.novo:":
                            val_antigo = ""

                        # busca o próximo "Val.novo:" dentro da mesma data
                        val_novo = ""
                        if val_antigo == "":
                            j = i + 1
                        else:
                            j = i + 2
                        while j < len(todos):
                            if interrompido:
                                break
                            prox_texto = todos[j]["texto"]
                            # se achou nova data, interrompe — o próximo par pertence à próxima data
                            if u.is_data(prox_texto):
                                break
                            if prox_texto == "Val.novo:" and j + 1 < len(todos):
                                if todos[j + 1]["texto"] != "5":
                                    val_novo = todos[j + 1]["texto"]
                                    if val_novo == motivo:
                                        motivo_encontrado = True
                                break
                            j += 1

                        linha_nova = {
                            "Instalacao": instalacao,
                            "Contrato": contrato,
                            "RE": re_atual,
                            "Data": data_atual,
                            "VAL.ANTIGO:": val_antigo or "",
                            "VAL.NOVO:": val_novo or "",
                        }
                        registros.append(linha_nova)

                        if motivo_encontrado:
                            break

                # Se não encontrou o motivo, faz o scroll e continua lendo
                if not motivo_encontrado:
                    scroll = session.findById("wnd[0]/usr").verticalScrollbar
                    if scroll.position >= scroll.maximum:
                        print_log(f"⚠️ Contrato {contrato}: final da tela atingido. Motivo não encontrado.")
                        registros = registros[-10:]
                        break
                    session.findById("wnd[0]").sendVKey(82)

            if registros:
                todos_registros.extend(registros)
                print_log(f"✅ Contrato {contrato} processado. | Restam {contratos_restantes} contratos")

            try:
                session.StartTransaction("ES21")
            except Exception:
                pass

    except Exception as e:
        print_log(f"❌ Ocorreu um erro: {e}")
        salvar_colheita(df_colheita, todos_registros, print_log)

    # Salva no final
    if todos_registros:
        salvar_colheita(df_colheita, todos_registros, print_log)
        print_log('🏁 Processamento finalizado. Resultados em "dados_coletados.xlsx".')

    try:
        session.StartTransaction("ES21")
        session.findById("wnd[0]").sendVKey(3)
    except Exception:
        pass

    
