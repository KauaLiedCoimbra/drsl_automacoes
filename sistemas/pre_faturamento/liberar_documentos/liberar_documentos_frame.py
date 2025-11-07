import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import style as s
from sistemas.pre_faturamento.liberar_documentos import liberar_documentos
import utils as u

def criar_frame_liberar_documentos(parent, btn_voltar=None):
    frame = ttk.Frame(parent, padding=10)
    frame.configure(style="Dracula.TFrame")

    if btn_voltar:
        btn_voltar.place(x=10, y=10)

    # ---------------------------
    # Logs
    # ---------------------------
    ttk.Label(frame, text="Logs de execução", font=("Segoe UI", 10, "bold"),
              foreground=s.DRACULA_TITLE, background=s.DRACULA_BG).pack(anchor="w")
    logs_widget = scrolledtext.ScrolledText(
        frame, height=12, bg=s.DRACULA_LOGS_WIDGET, fg=s.DRACULA_FG, insertbackground=s.DRACULA_FG
    )
    logs_widget.pack(fill="both", expand=True, pady=10)

    # ---------------------------
    # Parâmetros
    # ---------------------------
    matricula_var = tk.StringVar(value="6328490")
    layout_var = tk.StringVar(value="//CHARLES")

    param_frame = ttk.Frame(frame, style="Dracula.TFrame")
    param_frame.pack(fill="x", pady=5)

    ttk.Label(param_frame, text="Matrícula:", style="Dracula.TLabel").grid(row=0, column=0, padx=5, sticky="w")
    ttk.Entry(param_frame, textvariable=matricula_var, style="Dracula.TEntry").grid(row=0, column=1, padx=5, sticky="w")

    ttk.Label(param_frame, text="Layout:", style="Dracula.TLabel").grid(row=1, column=0, padx=5, sticky="w")
    ttk.Entry(param_frame, textvariable=layout_var, style="Dracula.TEntry").grid(row=1, column=1, padx=5, sticky="w")

    # ---------------------------
    # Funções internas
    # ---------------------------
    def executar():
        def rodar():
            try:
                liberar_documentos.executar_liberar_documentos(
                    logs_widget,
                    matricula=matricula_var.get(),
                    layout=layout_var.get()
                )
            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro na execução: {e}")

        threading.Thread(target=rodar, daemon=True).start()

    # ---------------------------
    # Botão Executar
    # ---------------------------
    ttk.Button(frame, text="Executar", command=executar, style="Dracula.TButton").pack(pady=10)

    return frame, logs_widget, None
