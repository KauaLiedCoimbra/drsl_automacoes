import tkinter as tk
from tkinter import ttk, messagebox
import threading
import utils as u
import style as s
from sistemas.notas_diarias.notas_diarias import executar_notas_diarias

def criar_frame_notas_diarias(parent, btn_voltar=None, destinatario_default="roberta.cardoso@cpfl.com.br"):
    frame = ttk.Frame(parent, padding=10, style="Custom.TFrame")

    interromper = tk.BooleanVar(value=False)

    # ---------------------------
    # ESTILOS
    # ---------------------------
    style = ttk.Style()
    style.configure("Custom.TFrame", background=s.DRACULA_BG)
    style.configure("Custom.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_FG)
    style.configure("Title.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_TITLE, font=("Segoe UI", 10, "bold"))
    style.configure("Custom.TButton", background=s.DRACULA_BUTTON_BG, foreground=s.DRACULA_FG)

    # ---------------------------
    # Logs
    # ---------------------------
    logs_widget = tk.Text(
        frame, height=16, width=80,
        state="disabled", bg=s.DRACULA_LOGS_WIDGET,
        fg=s.DRACULA_FG, relief="flat", wrap="word",
        insertbackground=s.DRACULA_FG
    )
    logs_widget.pack(fill="both", expand=True, pady=(0, 10))

    # ---------------------------
    # Campo de destinatário
    # ---------------------------
    destinatario_var = tk.StringVar(value=destinatario_default)

    email_frame = ttk.Frame(frame, style="Custom.TFrame")
    email_frame.pack(fill="x", pady=(0, 10))

    ttk.Label(email_frame, text="Destinatário do e-mail:", style="Custom.TLabel").pack(side="left")
    entrada_email = ttk.Entry(email_frame, textvariable=destinatario_var, width=40)
    entrada_email.pack(side="left", padx=(5,0))

    # ---------------------------
    # Botões
    # ---------------------------
    botoes_frame = ttk.Frame(frame, style="Custom.TFrame")
    botoes_frame.pack(pady=10)

    btn_executar = tk.Button(
        botoes_frame, text="▶ Enviar Notas",
        bg=s.DRACULA_BUTTON_BG, fg=s.DRACULA_FG,
        activebackground=s.DRACULA_BUTTON_ACTIVE,
        command=lambda: executar()
    )
    btn_executar.pack(side="left", padx=5)

    btn_interromper = tk.Button(
        botoes_frame, text="⛔ Interromper",
        bg="#ff5555", fg=s.DRACULA_FG,
        activebackground="#ff6e6e",
        command=lambda: interromper.set(True)
    )
    btn_interromper.pack(side="left", padx=5)

    # ---------------------------
    # Função executar()
    # ---------------------------
    def executar():
        destinatario = destinatario_var.get().strip()
        if not destinatario:
            messagebox.showwarning("Aviso", "Informe o destinatário do e-mail.")
            return

        interromper.set(False)

        def tarefa():
            try:
                executar_notas_diarias(
                    destinatario=destinatario,
                    logs_widget=logs_widget,
                    interromper_var=interromper
                )
                if not interromper.get():
                    u.print_log(logs_widget, "✅ E-mail enviado com sucesso!")
                else:
                    u.print_log(logs_widget, "⚠️ Operação interrompida pelo usuário.")
            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro: {e}")

        threading.Thread(target=tarefa, daemon=True).start()

    return frame, logs_widget, interromper
