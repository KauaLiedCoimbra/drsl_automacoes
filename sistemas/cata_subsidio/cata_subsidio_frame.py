import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import threading
import utils as u
import style as s

def criar_frame_cata_subsidio(parent, btn_voltar=None):
    frame = ttk.Frame(parent, padding=10, style="Custom.TFrame")

    # Variáveis
    periodo_inicio_var = tk.StringVar()
    periodo_fim_var = tk.StringVar()
    interromper = tk.BooleanVar(value=False)
    instalacoes = []

    info_labels = [
        "Histórico de consumo",
        "Faturas / Pagamentos / Parcelamento",
        "Devoluções de créditos",
        "Negativações / Protesto",
        "Faturas em PDF",
        "Consulta de leituras",
        "Informações de contrato"
    ]
    info_vars = {label: tk.BooleanVar(value=True) for label in info_labels}

    # ---------------------------
    # ESTILOS
    # ---------------------------
    style = ttk.Style()
    style.configure("Custom.TFrame", background=s.DRACULA_BG)
    style.configure("Custom.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_FG)
    style.configure("Title.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_TITLE, font=("Segoe UI", 10, "bold"))
    style.configure("Custom.TCheckbutton", background=s.DRACULA_BG, foreground=s.DRACULA_FG)
    style.configure("Custom.TButton", background=s.DRACULA_BUTTON_BG, foreground=s.DRACULA_FG)

    # ---------------------------
    # Linha de cima: Logs + Informações desejadas
    # ---------------------------
    top_frame = ttk.Frame(frame, style="Custom.TFrame")
    top_frame.pack(fill="both", expand=True, pady=(0, 10))

    # Logs à esquerda
    logs_widget = tk.Text(
        top_frame, height=14, width=60,
        state="disabled", bg=s.DRACULA_LOGS_WIDGET,
        fg=s.DRACULA_FG, relief="flat", wrap="word",
        insertbackground=s.DRACULA_FG
    )
    logs_widget.pack(side="left", fill="both", expand=True, padx=(0, 10))

    # Informações desejadas à direita (vertical)
    direito_frame = ttk.Frame(top_frame, style="Custom.TFrame")
    direito_frame.pack(side="left", fill="y", expand=False)

    ttk.Label(
        direito_frame, text="Informações desejadas:",
        style="Title.TLabel"
    ).pack(anchor="w", pady=(0, 5))

    for label in info_labels:
        ttk.Checkbutton(
            direito_frame,
            text=label,
            variable=info_vars[label],
            style="Custom.TCheckbutton"
        ).pack(anchor="w", pady=2)

    # ---------------------------
    # Período
    # ---------------------------
    periodo_frame = ttk.Frame(frame, style="Custom.TFrame")
    periodo_frame.pack(fill="x", pady=(5, 10))

    ttk.Label(periodo_frame, text="Período:", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
    ttk.Label(periodo_frame, text="Início:", style="Custom.TLabel").grid(row=0, column=1, sticky="e")

    date_inicio = DateEntry(
        periodo_frame, textvariable=periodo_inicio_var, width=12,
        background=s.DRACULA_WIDGET_BG, foreground=s.DRACULA_FG,
        borderwidth=1, relief="flat", date_pattern="dd/mm/yyyy"
    )
    date_inicio.configure(selectbackground=s.DRACULA_BUTTON_BG)
    date_inicio.grid(row=0, column=2, padx=5)

    ttk.Label(periodo_frame, text="Fim:", style="Custom.TLabel").grid(row=0, column=3, sticky="e")

    date_fim = DateEntry(
        periodo_frame, textvariable=periodo_fim_var, width=12,
        background=s.DRACULA_WIDGET_BG, foreground=s.DRACULA_FG,
        borderwidth=1, relief="flat", date_pattern="dd/mm/yyyy"
    )
    date_fim.configure(selectbackground=s.DRACULA_BUTTON_BG)
    date_fim.grid(row=0, column=4, padx=5)

    # ---------------------------
    # Instalações (tags)
    # ---------------------------
    ttk.Label(frame, text="Instalações:", style="Title.TLabel").pack(anchor="w", pady=(5, 3))
    instalacoes_frame = ttk.Frame(frame, style="Custom.TFrame")
    instalacoes_frame.pack(fill="x")

    instalacoes_tags_frame = ttk.Frame(instalacoes_frame, style="Custom.TFrame")
    instalacoes_tags_frame.pack(fill="x", pady=(0, 5))

    entrada_instalacao = tk.Entry(
        instalacoes_frame, width=10,
        bg=s.DRACULA_WIDGET_BG, fg=s.DRACULA_FG,
        insertbackground=s.DRACULA_FG,
        relief="flat", highlightthickness=1, highlightbackground=s.DRACULA_BUTTON_BG
    )
    entrada_instalacao.pack(fill="x")

    def adicionar_instalacao(event=None):
        valor = entrada_instalacao.get().strip()
        if not valor:
            return

        if valor in instalacoes:
            u.print_log(logs_widget, f"⚠️ Instalação '{valor}' já adicionada.")
            entrada_instalacao.delete(0, "end")
            return

        if len(instalacoes) >= 10:
            u.print_log(logs_widget, "⚠️ Limite de 10 instalações atingido.")
            entrada_instalacao.delete(0, "end")
            return

        instalacoes.append(valor)

        tag = tk.Frame(instalacoes_tags_frame, bg=s.DRACULA_WIDGET_BG, bd=1, relief="ridge")
        tag.pack(side="left", padx=3, pady=3)

        lbl_text = tk.Label(tag, text=valor, bg=s.DRACULA_WIDGET_BG, fg=s.DRACULA_FG, padx=6)
        lbl_text.pack(side="left")

        btn_x = tk.Label(tag, text="✕", bg=s.DRACULA_WIDGET_BG, fg="#ff5555", padx=4, cursor="hand2", font=("Segoe UI", 9, "bold"))
        btn_x.pack(side="left")
        btn_x.bind("<Enter>", lambda e: btn_x.config(fg="#ff6e6e"))
        btn_x.bind("<Leave>", lambda e: btn_x.config(fg="#ff5555"))
        btn_x.bind("<Button-1>", lambda e, v=valor, f=tag: remover_instalacao(v, f))

        entrada_instalacao.delete(0, "end")

    def remover_instalacao(valor, frame_ref):
        instalacoes.remove(valor)
        frame_ref.destroy()

    entrada_instalacao.bind("<Return>", adicionar_instalacao)

    # ---------------------------
    # Botões
    # ---------------------------
    botoes_frame = ttk.Frame(frame, style="Custom.TFrame")
    botoes_frame.pack(pady=10)

    btn_executar = tk.Button(
        botoes_frame, text="▶ Executar",
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
        if not instalacoes:
            messagebox.showwarning("Aviso", "Informe ao menos uma instalação.")
            return

        infos_selecionadas = [k for k, v in info_vars.items() if v.get()]
        if not infos_selecionadas:
            messagebox.showwarning("Aviso", "Selecione ao menos uma informação.")
            return

        inicio = periodo_inicio_var.get()
        fim = periodo_fim_var.get()
        interromper.set(False)

        from sistemas.cata_subsidio.cata_subsidio import coletar_dados

        def tarefa():
            try:
                coletar_dados(
                    instalacoes,
                    infos_selecionadas,
                    periodo_inicio=inicio,
                    periodo_fim=fim,
                    logs_widget=logs_widget,
                    interromper_var=interromper
                )
                if not interromper.get():
                    u.print_log(logs_widget, "✅ Consulta concluída!")
                else:
                    u.print_log(logs_widget, "⚠️ Consulta interrompida pelo usuário.")
            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro: {e}")

        threading.Thread(target=tarefa, daemon=True).start()

    return frame, logs_widget, interromper
