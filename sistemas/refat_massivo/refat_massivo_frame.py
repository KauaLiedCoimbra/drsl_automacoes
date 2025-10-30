import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from sistemas.refat_massivo.refat_massivo import executar_refat_massivo
import threading
import style as s

def criar_frame_refat_massivo(parent, btn_voltar=None):
    frame = ttk.Frame(parent, padding=10)
    frame.configure(style="Dracula.TFrame")

    if btn_voltar:
        btn_voltar.place(x=10, y=10)

    # ---------------------------
    # Estilos ttk
    # ---------------------------
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Dracula.TFrame", background=s.DRACULA_BG)
    style.configure("Dracula.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_FG, font=("Segoe UI", 10))
    style.configure("DraculaBold.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_TITLE, font=("Segoe UI", 10, "bold"))
    style.configure("Dracula.TEntry", fieldbackground=s.DRACULA_WIDGET_BG, foreground=s.DRACULA_FG)
    style.configure("Dracula.TButton", background=s.DRACULA_BUTTON_BG, foreground=s.DRACULA_FG, relief="flat")
    style.map("Dracula.TButton", background=[("active", s.DRACULA_BUTTON_ACTIVE)])

    # ---------------------------
    # Logs
    # ---------------------------
    ttk.Label(frame, text="Logs de execução", style="DraculaBold.TLabel").pack(anchor="w")
    logs_widget = scrolledtext.ScrolledText(
        frame,
        height=12,
        bg=s.DRACULA_LOGS_WIDGET,
        fg=s.DRACULA_FG,
        insertbackground=s.DRACULA_FG,
        relief="flat",
        highlightbackground=s.DRACULA_WIDGET_BG,
        highlightthickness=1,
        wrap="word",
    )
    logs_widget.pack(fill="both", expand=True, pady=10)

    # ---------------------------
    # Variáveis
    # ---------------------------
    caminho_planilha_var = tk.StringVar()
    local_var = tk.StringVar(value="/interf")
    periodo_var = tk.StringVar(value="2025/10")
    tamanho_lote_var = tk.IntVar(value=2500)
    coluna_var = tk.StringVar(value="Instalação")

    # Flag de interrupção
    interromper_execucao = tk.BooleanVar(value=False)

    # ---------------------------
    # Funções internas
    # ---------------------------
    def selecionar_planilha():
        caminho = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if caminho:
            caminho_planilha_var.set(caminho)

    def log(msg):
        logs_widget.insert(tk.END, msg + "\n")
        logs_widget.see(tk.END)
        logs_widget.update()

    def executar():
        if not caminho_planilha_var.get():
            log("⚠️ Selecione uma planilha!")
            return

        interromper_execucao.set(False)  # Resetar flag ao iniciar

        def rodar():
            log("🔹 Iniciando processamento...")
            try:
                # Aqui passamos a flag de interrupção para a função de processamento
                caminho_resultado = executar_refat_massivo(
                    logs_widget,
                    caminho_planilha=caminho_planilha_var.get(),
                    periodo=periodo_var.get(),
                    tamanho_lote=tamanho_lote_var.get(),
                    p_file=local_var.get(),
                    coluna=coluna_var.get(),
                    interromper_flag=interromper_execucao
                )
                log(f"✅ Processamento finalizado! Arquivo salvo em:\n{caminho_resultado}")
            except Exception as e:
                log(f"⚠️ Erro: {e}")

        threading.Thread(target=rodar, daemon=True).start()

    def interromper():
        log("⏹ Interrupção solicitada. Salvando progresso...")
        interromper_execucao.set(True)

    # ---------------------------
    # Layout do frame
    # ---------------------------
    input_frame = ttk.Frame(frame, style="Dracula.TFrame")
    input_frame.pack(pady=5, fill="x")

    # Linha 0: Planilha
    ttk.Label(input_frame, text="Planilha:", style="Dracula.TLabel").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=caminho_planilha_var, width=40, style="Dracula.TEntry").grid(row=0, column=1, columnspan=2, sticky="w")
    ttk.Button(input_frame, text="Selecionar", command=selecionar_planilha, style="Dracula.TButton").grid(row=0, column=3, padx=5, pady=5)

    # Linha 1: Período e Local
    ttk.Label(input_frame, text="Período:", style="Dracula.TLabel").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=periodo_var, style="Dracula.TEntry").grid(row=1, column=1, sticky="w", padx=5, pady=5)
    ttk.Label(input_frame, text="Local:", style="Dracula.TLabel").grid(row=1, column=2, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=local_var, style="Dracula.TEntry").grid(row=1, column=3, sticky="w", padx=5, pady=5)

    # Linha 2: Tamanho do Lote e Coluna
    ttk.Label(input_frame, text="Tamanho do Lote:", style="Dracula.TLabel").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=tamanho_lote_var, style="Dracula.TEntry").grid(row=2, column=1, sticky="w", padx=5, pady=5)
    ttk.Label(input_frame, text="Coluna:", style="Dracula.TLabel").grid(row=2, column=2, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=coluna_var, style="Dracula.TEntry").grid(row=2, column=3, sticky="w", padx=5, pady=5)

    # Botões Executar e Interromper
    botao_frame = ttk.Frame(frame, style="Dracula.TFrame")
    botao_frame.pack(pady=10)
    ttk.Button(botao_frame, text="Executar", command=executar, style="Dracula.TButton").pack(side="left", padx=5)
    ttk.Button(botao_frame, text="Interromper", command=interromper, style="Dracula.TButton").pack(side="left", padx=5)

    return frame, logs_widget, interromper_execucao
