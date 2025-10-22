from tkinter import ttk, scrolledtext
import threading
import mapear_sap.mapear_sap as mapear_sap  # módulo que você deve ter para executar a lógica
import utils as u
import style

def criar_frame_sap_map(parent, btn_voltar=None):
    """Cria o frame completo do sistema de mapeamento SAP"""
    frame = ttk.Frame(parent, padding=10)
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

    def executar_mapeamento_thread():
        def target():
            try:
                mapear_sap.executar_mapeamento(lambda msg: u.print_log(logs_widget, msg))
            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro durante execução: {e}")

        threading.Thread(target=target, daemon=True).start()

    # Botão de execução
    ttk.Button(btn_frame, text="▶ Executar Mapeamento SAP", command=executar_mapeamento_thread).pack(side="left", padx=5, ipady=5)

    style.aplicar_estilo(frame)

    return frame, logs_widget, None