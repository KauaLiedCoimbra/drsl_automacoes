import tkinter as tk
from tkinter import ttk
from es21_frame import criar_frame_es21

# ---------------------------
# Dados iniciais
# ---------------------------
nucleos = {
    "Administrativo": [],
    "Qualidade": [],
    "Pré-Faturamento": [],
    "Pós-Faturamento": ["Logs de bloqueio - ES21"],  # Botão adaptável
    "Reclamação": [],
    "Jurídico": []
}
sistemas_frames = {
    "Logs de bloqueio - ES21": criar_frame_es21,
}
frames_criados = {}
# ---------------------------
# Janela principal
# ---------------------------
root = tk.Tk()
root.title("Automações do Kauã")
root.geometry("700x500")
root.resizable(False, False)

# Dicionário com conteúdos específicos de sistemas
sistemas_frames = {
    "Logs de bloqueio - ES21": criar_frame_es21,
}
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
        canvas = tk.Canvas(systems_container, height=200, highlightthickness=0)
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

    ttk.Label(system_frame, text=f"Sistema: {sistema}", font=("Arial", 16, "bold")).pack(pady=20)

    if sistema in sistemas_frames:
        frame, logs_widget, interromper = sistemas_frames[sistema](system_frame, btn_voltar=btn_voltar)
        frames_criados[sistema] = (frame, logs_widget, interromper)
        btn_voltar.place(x=10, y=10)
        frame.pack(fill="both", expand=True)
    else:
        ttk.Label(system_frame, text="Conteúdo do sistema aqui (vazio por enquanto)").pack(pady=10)

    system_frame.pack(fill="both", expand=True)

def voltar_para_nucleos():
    system_frame.pack_forget()
    frame_nucleos.pack(fill="x")
    systems_container.pack(fill="both", expand=True)
    btn_voltar.place_forget()

# ---------------------------
# Botão Voltar fixo (persistente)
# ---------------------------
btn_voltar = ttk.Button(root, text="🔙 Voltar", command=lambda: voltar_para_nucleos())
btn_voltar.place(x=10, y=10)   # posição fixa no canto superior esquerdo
btn_voltar.place_forget()      # começa escondido

# ---------------------------
# Títulos e botões dos núcleos
# ---------------------------
ttk.Label(frame_nucleos, text="Automações do Kauã", font=("Arial", 18, "bold")).grid(row=0, column=0, columnspan=3, pady=(0,10))
ttk.Label(frame_nucleos, text="Escolha o núcleo:", font=("Arial", 14, "bold")).grid(row=1, column=0, columnspan=3, pady=(0,10))

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
root.mainloop()