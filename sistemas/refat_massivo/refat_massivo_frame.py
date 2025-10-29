import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from sistemas.refat_massivo.refat_massivo import executar_refat_massivo
import threading

def criar_frame_refat_massivo(parent, btn_voltar=None):
    frame = ttk.Frame(parent, padding=10)
    if btn_voltar:
        btn_voltar.place(x=10, y=10)

    logs_widget = scrolledtext.ScrolledText(frame, height=12)
    logs_widget.pack(fill="both", expand=True, pady=10)

    # Variáveis de input
    caminho_planilha_var = tk.StringVar()
    local_var = tk.StringVar(value="/interf")
    periodo_var = tk.StringVar(value="2025/10")
    tamanho_lote_var = tk.IntVar(value=2500)
    coluna_var = tk.StringVar(value="Instalação")

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

        def rodar():
            log("🔹 Iniciando processamento...")
            try:
                relatorios = executar_refat_massivo(
                    logs_widget,
                    caminho_planilha=caminho_planilha_var.get(),
                    periodo=periodo_var.get(),
                    tamanho_lote=tamanho_lote_var.get(),
                    p_file=local_var.get(),
                    coluna=coluna_var.get(),
                )
                log(f"✅ Processamento finalizado! {len(relatorios)} arquivos gerados.")
            except Exception as e:
                log(f"⚠️ Erro: {e}")

        threading.Thread(target=rodar, daemon=True).start()

    # ---------------------------
    # Layout do frame
    # ---------------------------
    input_frame = ttk.Frame(frame)
    input_frame.pack(pady=5, fill="x")

    # Linha 0: Planilha
    ttk.Label(input_frame, text="Planilha:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=caminho_planilha_var, width=40).grid(row=0, column=1, columnspan=2, sticky="w")
    ttk.Button(input_frame, text="Selecionar", command=selecionar_planilha).grid(row=0, column=3, padx=5, pady=5)

    # Linha 1: Período e Local
    ttk.Label(input_frame, text="Período:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=periodo_var).grid(row=1, column=1, sticky="w", padx=5, pady=5)

    ttk.Label(input_frame, text="Local:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=local_var).grid(row=1, column=3, sticky="w", padx=5, pady=5)

    # Linha 2: Tamanho do Lote e coluna
    ttk.Label(input_frame, text="Tamanho do Lote:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=tamanho_lote_var).grid(row=2, column=1, sticky="w", padx=5, pady=5)

    ttk.Label(input_frame, text="Tamanho do Lote:").grid(row=2, column=2, sticky="w", padx=5, pady=5)
    ttk.Entry(input_frame, textvariable=coluna_var).grid(row=2, column=3, sticky="w", padx=5, pady=5)

    # Botão Executar
    ttk.Button(frame, text="Executar", command=executar).pack(pady=10)

    interromper = False
    return frame, logs_widget, interromper
