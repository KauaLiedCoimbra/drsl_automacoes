import re
import psutil
import time

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

def print_log(widget, msg):
    if not widget:
        return

    def _update():
        widget.config(state="normal")
        _, bottom = widget.yview()
        no_final = bottom >= 0.9  # perto do final

        widget.insert("end", msg + "\n")
        if no_final:
            widget.see("end")

        widget.config(state="disabled")

    widget.after(0, _update)  # garante execução na thread principal

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