import win32com.client
import pandas as pd
import re
import signal
import sys

# --- FLAG GLOBAL PARA INTERRUPÇÃO ---
interrompido = False

def salvar_e_sair(signum, frame):
    global interrompido
    interrompido = True
    print("\n⚠ Execução interrompida! Salvando dados coletados até agora...")
    salvar_colheita(df_colheita, todos_registros)
    session.StartTransaction("ES21")
    sys.exit(0)

def salvar_colheita(df_colheita, todos_registros):
    if not todos_registros:
        return
    df_colheita_save = pd.concat([df_colheita, pd.DataFrame(todos_registros)], ignore_index=True)
    try:
        with pd.ExcelWriter("2_registros_coletados.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_colheita_save.to_excel(writer, sheet_name="Coleta", index=False)
        print("✅ Dados salvos em '2_registros_coletados.xlsx'")
    except FileNotFoundError:
        df_colheita_save.to_excel("2_registros_coletados.xlsx", index=False)
        print("✅ Arquivo '2_registros_coletados.xlsx' criado do zero.")

def is_data(data):
    padrao_data = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    if padrao_data.match(data):
        return True
    else:
        return False

signal.signal(signal.SIGINT, salvar_e_sair)

# Lê planilhas
df = pd.read_excel("1_banco_contratos.xlsx")
try:
    df_colheita = pd.read_excel("2_registros_coletados.xlsx")
except FileNotFoundError:
    df_colheita = pd.DataFrame(columns=['Instalacao','Contrato','RE','Data','VAL.ANTIGO:','VAL.NOVO:'])

# Corrige colunas
for col in ['INSTALACAO','CONTRATOS', 'MOTIVO']:
    if col in df.columns:
        df[col] = df[col].apply(lambda x: str(int(x)) if isinstance(x, float) else str(x)).str.strip()
    else:
        raise ValueError(f"Coluna {col} não encontrada na planilha.")

# Conexão SAP
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Maximiza e acessa ES21
session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "es21"
session.findById("wnd[0]").sendVKey(0)

todos_registros = []
total_contratos = len(df)

try:
    for index, row in df.iterrows():
        instalacao = row['INSTALACAO']
        contrato = row['CONTRATOS']
        motivo = row["MOTIVO"].zfill(2)
        contratos_restantes = total_contratos - (index + 1)
        print(f'🔍 Processando contrato {contrato}... Motivo: {motivo}')

        if interrompido:
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
                texto = elem["texto"]

                # 🟢 Quando encontrar uma nova data, atualiza o RE e a data em uso
                if is_data(texto) and not (i > 0 and todos[i - 1]["texto"] in ["Val.antigo:", "Val.novo:"]):
                    if texto != data_atual:
                        data_atual = texto
                        re_atual = todos[i - 1]["texto"] if i > 0 else ""
                    
                 # 🔵 captura pares de valores dentro da mesma data
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
                        prox_texto = todos[j]["texto"]
                        # se achou nova data, interrompe — o próximo par pertence à próxima data
                        if is_data(prox_texto):
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
                    print(f"⚠️ Contrato {contrato}: final da tela atingido. Motivo não encontrado.")
                    registros = registros[-10:]  # mantém os últimos 10 registros para referência
                    break
                session.findById("wnd[0]").sendVKey(82)

        if registros:
            todos_registros.extend(registros)
            print(f"✅ Contrato {contrato} processado. | Restam {contratos_restantes} contratos")

        session.StartTransaction("ES21")

except Exception as e:
    print(f"❌ Ocorreu um erro: {e}")
    salvar_colheita(df_colheita, todos_registros)

# Salva no final
if todos_registros:
    salvar_colheita(df_colheita, todos_registros)

session.StartTransaction("ES21")
session.findById("wnd[0]").sendVKey(3)
print('🏁 Processamento finalizado. Resultados em "2_registros_coletados.xlsx".')
