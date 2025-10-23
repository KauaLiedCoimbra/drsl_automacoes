from tkinter import ttk, scrolledtext, filedialog
import threading
from logs_bloqueio import logs_bloqueio
import utils as u
import style
import pandas as pd

def criar_frame_logs_bloqueio(parent, btn_voltar=None):
    """Cria o frame completo do ES21 com logs e botões"""
    frame = ttk.Frame(parent, padding=10)
    if btn_voltar:
        btn_voltar.place(x=10, y=10)

    # ScrolledText para logs
    logs_widget = scrolledtext.ScrolledText(
    frame,
    width=90,
    height=15,
    font=("Consolas", 10),  # fundo do Dracula
    fg=style.DRACULA_FG,
    bg=style.DRACULA_LOGS_WIDGET,  # cor do cursor (mesmo que não apareça)
    relief="flat",             # sem bordas 3D
    borderwidth=5,
    padx=5,
    pady=5,
)
    logs_widget.pack(fill="both", expand=True)
    logs_widget.config(state="disabled")

    # Frame para botões
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=(10,0))

    caminho_planilha = None
    df_resultado = pd.DataFrame()

    def anexar_planilha():
        nonlocal caminho_planilha
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if caminho:
            caminho_planilha = caminho
            u.print_log(logs_widget, f"📄 Arquivo selecionado: {caminho}")
        else:
            u.print_log(logs_widget, "⚠ Nenhum arquivo selecionado")

    def executar_logs_bloqueio_thread():
        nonlocal caminho_planilha
        if not caminho_planilha:
            u.print_log(logs_widget, "❌ Nenhum arquivo Excel selecionado.")
            return
        
        def target():
            nonlocal df_resultado
            try:
                df_resultado = logs_bloqueio.executar_logs_bloqueio(caminho_planilha, lambda msg: u.print_log(logs_widget, msg))
                u.print_log(logs_widget, "✅ Execução concluída. Dados prontos para download.")
            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro durante execução: {e}")
            finally:
                logs_bloqueio.interrompido = False

        threading.Thread(target=target, daemon=True).start()

    def interromper():
        logs_bloqueio.interrompido = True
        u.print_log(logs_widget, "⚠ Interrupção solicitada pelo usuário")

    def baixar_arquivo():
            nonlocal df_resultado
            if df_resultado.empty:
                u.print_log(logs_widget, "❌ Nenhum dado disponível para download.")
                return

            caminho = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Salvar planilha logs_bloqueio"
            )
            if caminho:
                df_resultado.to_excel(caminho, index=False)
                u.print_log(logs_widget, f"✅ Arquivo salvo em: {caminho}")
   
    # Botões
    ttk.Button(btn_frame, text="📎 Anexar planilha", command=anexar_planilha).pack(side="left", padx=5, ipady=5)
    ttk.Button(btn_frame, text="▶ Executar", command=executar_logs_bloqueio_thread).pack(side="left", padx=5, ipady=5)
    ttk.Button(btn_frame, text="⏹ interromper", command=interromper).pack(side="left", padx=5, ipady=5)
    ttk.Button(btn_frame, text="💾 Baixar resultados", command=baixar_arquivo).pack(side="left", padx=5, ipady=5)

    style.aplicar_estilo(frame)

    return frame, logs_widget, interromper