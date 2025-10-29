import tkinter as tk
from tkinter import ttk
from sistemas.logs_bloqueio.logs_bloqueio_frame import criar_frame_logs_bloqueio
from sistemas.mapear_sap.mapear_sap_frame import criar_frame_sap_map
from sistemas.cata_erro.cata_erro_frame import criar_frame_cata_erro
from sistemas.converte_parquet.conversor_parquet import criar_frame_conversor_parquet
import style
import ctypes
import os
import sys

# ---------------------------
# Dados iniciais
# ---------------------------
nucleos = {
    "Administrativo": [],
    "Qualidade": ["Mapeamento SAP", "Conversor Parquet"],
    "Pré-Faturamento": [],
    "Pós-Faturamento": ["Logs de bloqueio", "Cata-erro"],
    "Reclamação": [],
    "Jurídico": ["Cata-subsídio"]
}
sistemas_frames = {
    "Logs de bloqueio": criar_frame_logs_bloqueio,
    "Mapeamento SAP": criar_frame_sap_map,
    "Cata-erro": criar_frame_cata_erro,
    "Conversor Parquet": criar_frame_conversor_parquet,
    "Cata-subsídio": None
}
frames_criados = {}
# ---------------------------
# Janela principal
# ---------------------------
root = tk.Tk()
root.title("Automações do Kauã")
ctypes.windll.shcore.SetProcessDpiAwareness(1)
scale_factor = root.winfo_fpixels('1i') / 72  # pixels por polegada / DPI base
root.tk.call('tk', 'scaling', 2.0)
# ---------------------------
# Tamanho dinâmico da janela
# ---------------------------
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calcula o tamanho da janela proporcional à tela
window_width = int(screen_width * 0.7)
window_height = int(screen_height * 0.75)

# Define limites para não ficar pequeno demais ou gigante
window_width = max(1000, min(window_width, 1600))
window_height = max(700, min(window_height, 1000))

# Centraliza a janela
x_pos = int((screen_width - window_width) / 2)
y_pos = int((screen_height - window_height) / 4)

root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
root.resizable(False, False)
# ---------------------------
# Ícone
# ---------------------------
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(".")

icon_path = os.path.join(base_path, "robotic-hand.ico")

# Ícone da janela principal
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)

# Ícone da barra de tarefas (Windows)
try:
    from ctypes import windll
    appid = u"kaua.automatron"  # identificador único (pode mudar o nome)
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
except Exception:
    pass
# ---------------------------
# Frames
# ---------------------------
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill="both", expand=True)

# Frame de núcleos
frame_nucleos = ttk.Frame(main_frame, padding=10)
frame_nucleos.pack(fill="x")

# Frame de sistemas
systems_container = ttk.Frame(main_frame, padding=10)
systems_container.pack(fill="both", expand=True)

# Frame para cada sistema individual
system_frame = ttk.Frame(main_frame, padding=10)

# ---------------------------
# Funções
# ---------------------------
def abrir_sistemas(nucleo):
    """Mostra os sistemas disponíveis para o núcleo selecionado."""
    # Limpa frame principal do sistema
    for widget in systems_container.winfo_children():
        widget.destroy()

    # Título do núcleo
    ttk.Label(systems_container, text=f"Sistemas do núcleo: {nucleo}",
              font=("Arial", 14, "bold")).pack(pady=10)

    sistemas = nucleos[nucleo]
    if sistemas:
        # Canvas rolável apenas se houver sistemas
        canvas = tk.Canvas(systems_container, height=200, bg=style.DRACULA_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(systems_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Adiciona os botões dentro do scrollable_frame
        for sistema in sistemas:
            ttk.Button(scrollable_frame, text=sistema,
                       command=lambda s=sistema: abrir_frame_sistema(s)).pack(pady=5, fill="x", anchor="center")
    else:
        # Se não houver sistemas, só mostra mensagem centralizada
        ttk.Label(systems_container, text="Nenhum sistema disponível").pack(pady=10)

def abrir_frame_sistema(sistema):
    """Abre um novo frame dentro da janela para o sistema selecionado."""
    frame_nucleos.pack_forget()
    systems_container.pack_forget()

    for widget in system_frame.winfo_children():
        widget.destroy()

    ttk.Label(system_frame, text=f"Sistema: {sistema}", font=("Consolas", 22, "bold"),
              foreground=style.DRACULA_TITLE, background=style.DRACULA_BG).pack(pady=30)

    if sistema in sistemas_frames:
        frame, logs_widget, interromper = sistemas_frames[sistema](system_frame, btn_voltar=btn_voltar)
        frames_criados[sistema] = (frame, logs_widget, interromper)
        btn_voltar.place(x=10, y=10)
        frame.pack(fill="both", expand=True)
    else:
        ttk.Label(system_frame, text="Conteúdo do sistema aqui (vazio por enquanto)",
                  font=("Consolas", 16), foreground=style.DRACULA_FG, background=style.DRACULA_BG).pack(pady=20)

    system_frame.pack(fill="both", expand=True)

def voltar_para_nucleos():
    system_frame.pack_forget()
    frame_nucleos.pack(fill="x")
    systems_container.pack(fill="both", expand=True)
    btn_voltar.place_forget()

# ---------------------------
# Botão Voltar fixo (persistente)
# ---------------------------
btn_voltar = ttk.Button(root, text="🔙 Voltar", command=lambda: voltar_para_nucleos(), width=12)
btn_voltar.place(x=10, y=10)   # posição fixa no canto superior esquerdo
btn_voltar.place_forget()      # começa escondido

# ---------------------------
# Títulos e botões dos núcleos
# ---------------------------
ttk.Label(frame_nucleos, text="DRSL AUTOMAÇÕES",
          font=("Consolas", 26, "bold"), foreground="#ff79c6", background=style.DRACULA_BG).grid(row=0, column=0, columnspan=3, pady=(10))
ttk.Label(frame_nucleos, text="Escolha o núcleo:",
          font=("Consolas", 20), foreground=style.DRACULA_FG, background=style.DRACULA_BG).grid(row=1, column=0, columnspan=3, pady=(10))

for i, nucleo in enumerate(nucleos.keys()):
    row = 2 + i // 3
    col = i % 3
    ttk.Button(frame_nucleos, text=nucleo, width=20,
               command=lambda n=nucleo: abrir_sistemas(n)).grid(row=row, column=col, padx=10, pady=5)

for col in range(3):
    frame_nucleos.grid_columnconfigure(col, weight=1)

# ---------------------------
# Inicializa interface
# ---------------------------
style.aplicar_estilo(root)
root.mainloop()