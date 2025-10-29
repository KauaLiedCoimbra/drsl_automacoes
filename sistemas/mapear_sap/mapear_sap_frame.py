from tkinter import ttk, scrolledtext, filedialog
from sistemas.mapear_sap import mapear_sap  # seu módulo SAP
import utils as u
import style

def criar_frame_sap_map(parent, btn_voltar=None):
    """Cria o frame completo do sistema de mapeamento SAP"""
    frame = ttk.Frame(parent, padding=10)
    if btn_voltar:
        btn_voltar.place(x=10, y=10)

    # Área de logs
    logs_widget = scrolledtext.ScrolledText(
        frame,
        width=90,
        height=15,
        font=("Consolas", 10),
        fg=style.DRACULA_FG,
        bg=style.DRACULA_LOGS_WIDGET,
        relief="flat",
        borderwidth=5,
        padx=5,
        pady=5,
    )
    logs_widget.pack(fill="both", expand=True)
    logs_widget.config(state="disabled")

    # Frame de botões
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=(10, 0))

    # Função para executar o mapeamento SAP em thread
    def executar_mapeamento_thread():
        mapear_sap.transcrever_sap_linear(lambda msg: u.print_log(logs_widget, msg))

    # Botão de execução
    ttk.Button(
        btn_frame,
        text="▶ Executar Mapeamento SAP",
        command=executar_mapeamento_thread
    ).pack(side="left", padx=5, ipady=5)

    # Função para baixar o arquivo gerado
    def baixar_arquivo():
        caminho = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Arquivos de texto", "*.txt")],
            title="Salvar arquivo SAP"
        )
        if caminho:
            conteudo = mapear_sap.obter_conteudo_gerado()  # pega o conteúdo em memória
            with open(caminho, "w", encoding="utf-8") as f:
                f.write(conteudo)
            u.print_log(logs_widget, f"✅ Arquivo salvo em: {caminho}")

    # Botão de download
    ttk.Button(
        btn_frame,
        text="⬇ Baixar Arquivo",
        command=baixar_arquivo
    ).pack(side="left", padx=5, ipady=5)

    style.aplicar_estilo(frame)

    return frame, logs_widget, None