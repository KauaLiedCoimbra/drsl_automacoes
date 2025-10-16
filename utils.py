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

def aguardar_carregamento_sap(session, timeout=10, intervalo=0.3):
    inicio = time.time()
    while True:
        try:
            # session.Busy é True enquanto SAP está processando
            if not session.Busy:
                return True
        except Exception:
            pass

        if time.time() - inicio > timeout:
            print_log(f"⚠️ SAP não respondeu em {timeout} segundos.")
            return False
        
        time.sleep(intervalo)