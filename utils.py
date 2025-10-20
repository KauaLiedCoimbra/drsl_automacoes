import re
import time

def is_data(data):
    padrao_data = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    if padrao_data.match(data):
        return True
    else:
        return False
    
def print_log(widget, msg):
    """Insere mensagem no ScrolledText e mantém scroll no final."""
    widget.config(state="normal")   # habilita escrita temporariamente
    widget.insert("end", msg + "\n")
    #widget.see("end")               # scroll automático para o final
    widget.config(state="disabled")

def normalizar_colunas(planilha):
    novas_colunas = []
    for col in planilha.columns:
        c = col.upper().strip().encode("ascii", errors="ignore").decode("utf-8")
        novas_colunas.append(c)
    planilha.columns = novas_colunas
    return planilha