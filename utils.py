import re
import psutil
import time
import win32com.client as win32
import inspect
import tkinter as tk

DATA_REGEX = r"^(\d{2})\.(\d{2})\.(\d{4})$"

def is_data(data):
    padrao_data = re.compile(DATA_REGEX)
    if padrao_data.match(data):
        return True
    else:
        return False
    
def extrair_ano(data):
    match = re.match(DATA_REGEX, data)
    if match:
        return match.group(3)  # captura o ano
    return None

def extrair_mes(data):
    match = re.match(DATA_REGEX, data)
    if match:
        return match.group(2)  # captura o mês
    return None

def extrair_dia(data):
    match = re.match(DATA_REGEX, data)
    if match:
        return match.group(1)  # captura o dia
    return None

def print_log(widget=None, msg=None):
    # Se não for widget válido, só imprime no console
    if not isinstance(widget, tk.Text):
        caller = inspect.stack()[1]
        filename = caller.filename
        lineno = caller.lineno
        funcname = caller.function
        print(f"[print_log] Widget inválido ou ausente! ({filename}:{funcname}:{lineno}) | msg={msg}")
        return

    if msg is None:
        caller = inspect.stack()[1]
        filename = caller.filename
        lineno = caller.lineno
        funcname = caller.function
        msg = f"[AVISO] print_log chamado sem mensagem! ({filename}:{funcname}:{lineno})"

    def _update():
        widget.config(state="normal")
        _, bottom = widget.yview()
        no_final = bottom >= 0.9
        widget.insert("end", msg + "\n")
        if bottom >= 0.9:
            widget.see("end")
        widget.config(state="disabled")

    widget.after(0, _update)

def normalizar_colunas(planilha):
    novas_colunas = []
    for col in planilha.columns:
        c = col.upper().strip().encode("ascii", errors="ignore").decode("utf-8")
        novas_colunas.append(c)
    planilha.columns = novas_colunas
    return planilha

def fechar_sap_forcadamente():
    """Mata todos os processos SAP GUI ativos."""
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'saplogon' in proc.info['name'].lower():
            proc.kill()
    time.sleep(2)  # espera o processo fechar

def conectar_sap():
    try:
        SapGuiAuto = win32.GetObject("SAPGUI")
    except Exception:
        print("❌ Não foi possível acessar o SAP GUI. Verifique se o SAP está aberto e com scripting habilitado (Alt+F12 → Opções → Scripting).")
        return
    
    application = SapGuiAuto.GetScriptingEngine
    if application is None or application.Children.Count == 0:
        print("❌ Nenhuma conexão SAP ativa encontrada. Abra o SAP e entre em uma sessão antes de executar.")
        return

    connection = application.Children(0)
    if connection.Children.Count == 0:
        print("❌ Nenhuma sessão aberta no SAP. Entre em um sistema (ex: SE16) e tente novamente.")
        return
    
    session = connection.Children(0)
    session.findById("wnd[0]").maximize()
    return session

def abrir_transacao(session, transacao):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"+transacao
    session.findById("wnd[0]").sendVKey(0)