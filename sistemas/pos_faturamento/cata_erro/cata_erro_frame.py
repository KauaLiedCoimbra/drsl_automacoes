from tkinter import ttk, scrolledtext, filedialog
import threading
import pandas as pd
import re
import utils as u
import style

def criar_frame_cata_erro(parent, btn_voltar=None):
    """Cria o frame do Cata-Erro com logs e botões"""
    frame = ttk.Frame(parent, padding=10)
    if btn_voltar:
        btn_voltar.place(x=10, y=10) 

    # ---------------------------
    # Widget de logs
    # ---------------------------
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
        pady=5
    )
    logs_widget.pack(fill="both", expand=True)
    logs_widget.config(state="disabled")

    # ---------------------------
    # Frame para botões
    # ---------------------------
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=(10,0))

    caminho_planilha = None  # Variável para armazenar o arquivo selecionado

    # ---------------------------
    # Função para anexar planilha
    # ---------------------------
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

    # ---------------------------
    # Função para processar Cata-Erro
    # ---------------------------
    def executar_cata_erro():
        nonlocal caminho_planilha
        if not caminho_planilha:
            u.print_log(logs_widget, "❌ Nenhum arquivo Excel selecionado.")
            return

        def target():
            try:
                # Linhas de ruído a ignorar
                trechos_para_remover = [
                    "OBS", "Início Criar conta", "Informação adicional", "Docs.que",
                    "Empr.:", "Operação (Empresa", "Energia Compensada Positiva",
                    "_________________________________________________________________________",
                    "Faturamento residencial", "Instalação(ões)",
                    "Erro interno:", "Erro durante leitura na tabela",
                    "No total", "Fim    Criar conta:"
                ]

                # Carrega o Excel
                df = pd.read_excel(caminho_planilha)
                coluna = df.columns[0].strip()

                # Regex para contas
                regex_conta = re.compile(r'\(conta:\s*(\d{12})\)')

                resultados = []
                conta_atual = None

                # Processa linhas de baixo para cima
                for linha in reversed(df[coluna].tolist()):
                    linha_str = str(linha)
                    if any(t.lower() in linha_str.lower() for t in trechos_para_remover):
                        continue
                    match_conta = regex_conta.search(linha_str)
                    if match_conta:
                        conta_atual = match_conta.group(1)
                        continue
                    if conta_atual:
                        resultados.append({'CC': conta_atual, 'ERRO': linha_str})

                resultados.reverse()
                df_final = pd.DataFrame(resultados)

                # Salva arquivo final
                caminho_saida = caminho_planilha.replace(".xlsx", "_processado.xlsx")
                df_final.to_excel(caminho_saida, index=False)

                u.print_log(logs_widget, f"✅ Arquivo '{caminho_saida}' criado com {len(df_final)} linhas!")

            except Exception as e:
                u.print_log(logs_widget, f"❌ Erro durante execução: {e}")

        # Thread para não travar a interface
        threading.Thread(target=target, daemon=True).start()

    # ---------------------------
    # Botões
    # ---------------------------
    ttk.Button(btn_frame, text="📎 Anexar planilha", command=anexar_planilha)\
        .pack(side="left", padx=5, ipady=5)
    ttk.Button(btn_frame, text="▶ Executar Cata-Erro", command=executar_cata_erro)\
        .pack(side="left", padx=5, ipady=5)

    # Aplica estilo
    style.aplicar_estilo(frame)

    return frame, logs_widget, None
