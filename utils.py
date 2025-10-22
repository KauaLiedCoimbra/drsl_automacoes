import re

def is_data(data):
    padrao_data = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    if padrao_data.match(data):
        return True
    else:
        return False
    
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