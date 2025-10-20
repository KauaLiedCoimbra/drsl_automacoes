from tkinter import ttk, scrolledtext, filedialog
import threading
import es21
import utils as u
import style

def criar_frame_es21(parent, btn_voltar=None):
    """Cria o frame completo do ES21 com logs e botões"""
    frame = ttk.Frame(parent, padding=10)
    btn_voltar.place(x=10, y=10) 

    logs_frame = ttk.Frame(frame)
    logs_frame.pack(fill="both", expand=True)

    # ScrolledText para logs
    logs_widget = scrolledtext.ScrolledText(
    frame,
    width=90,
    height=15,
    font=("Consolas", 10),  # fundo do Dracula
    fg=style.DRACULA_FG,
    bg=style.DRACULA_BG,  # cor do cursor (mesmo que não apareça)
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

    def anexar_planilha():
        nonlocal caminho_planilha
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if caminho:
            caminho_planilha = caminho
            logs_widget.config(state="normal")
            logs_widget.insert("end", f"📄 Arquivo selecionado: {caminho}\n")
            logs_widget.config(state="disabled")
        #    logs_widget.see("end")
        else:
            logs_widget.config(state="normal")
            logs_widget.insert("end", "⚠ Nenhum arquivo selecionado\n")
            logs_widget.config(state="disabled")
        #    logs_widget.see("end")

    def executar_es21_thread():
        nonlocal caminho_planilha
        if not caminho_planilha:
            u.print_log(logs_widget, "❌ Nenhum arquivo Excel selecionado.")
            return
        
        def target():
            try:
                es21.executar_es21(caminho_planilha, lambda msg: u.print_log(logs_widget, msg))
            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro durante execução: {e}")
            finally:
                es21.interrompido = False  # reset após execução

        threading.Thread(target=target, daemon=True).start()

    def interromper():
        es21.interrompido = True
        u.print_log(logs_widget, "⚠ Interrupção solicitada pelo usuário")

    # Botões
    ttk.Button(btn_frame, text="📎 Anexar planilha", command=anexar_planilha).pack(side="left", padx=5, ipady=5)
    ttk.Button(btn_frame, text="▶ Executar ES21", command=executar_es21_thread).pack(side="left", padx=5, ipady=5)
    ttk.Button(btn_frame, text="⏹ interromper", command=interromper).pack(side="left", padx=5, ipady=5)

    style.aplicar_estilo(frame)

    return frame, logs_widget, interromper