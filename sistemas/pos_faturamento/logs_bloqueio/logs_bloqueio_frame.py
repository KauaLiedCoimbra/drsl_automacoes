from tkinter import ttk, filedialog
import tkinter as tk
import threading
from sistemas.pos_faturamento.logs_bloqueio import logs_bloqueio
import utils as u
import style
import pandas as pd
import time

def criar_frame_logs_bloqueio(parent, btn_voltar=None):
    """Cria o frame completo do ES21 com logs e botões"""
    frame = ttk.Frame(parent, padding=10)
    if btn_voltar:
        btn_voltar.place(x=10, y=10)

    # Frame interno para logs + barra
    logs_frame = ttk.Frame(frame)
    logs_frame.pack(fill="both", expand=True, pady=(0,5))

    # Barra de progresso estilizada
    style_pb = ttk.Style()
    style_pb.theme_use('clam')
    style_pb.configure(
        "Dracula.Horizontal.TProgressbar",
        troughcolor="#282a36",      # fundo
        background="#50fa7b",       # preenchimento
        thickness=20,               # altura
        bordercolor="#282a36",
        lightcolor="#50fa7b",
        darkcolor="#50fa7b"
    )
    progress = ttk.Progressbar(
        logs_frame, orient="horizontal", length=600, mode="determinate",
        style="Dracula.Horizontal.TProgressbar"
    )
    progress.pack(fill="x", pady=(0,5))

    # Label para mostrar % e ETA
    progress_label = ttk.Label(
        logs_frame, text="0% | ETA: --:--", foreground="#f8f8f2",
        background="#282a36", font=("Consolas", 10, "bold")
    )
    progress_label.pack(anchor="center", pady=(0,5))  # Coloca acima da barra, alinhada ao centro

    # ScrolledText para logs
    logs_widget = style.criar_logs_widget(frame)

    # Frame para botões
    btn_frame = ttk.Frame(frame)
    btn_frame.pack(fill="x", pady=(10,0))

    caminho_planilha = None
    df_resultado = pd.DataFrame()

    start_time = None
    contratos_processados = 0
    total_contratos = 0

    def iniciar_barra(max_value):
        nonlocal start_time, contratos_processados, total_contratos
        progress["maximum"] = max_value
        progress["value"] = 0
        contratos_processados = 0
        total_contratos = max_value
        start_time = time.time()
        progress_label.config(text="0% | ETA: --:--")

    def atualizar_barra(passo=1):
        nonlocal contratos_processados
        progress.step(passo)
        contratos_processados += passo

        percent = int((progress["value"] / progress["maximum"]) * 100)

        # Estimativa de tempo restante
        elapsed = time.time() - start_time
        if contratos_processados > 0:
            tempo_restante = elapsed / contratos_processados * (total_contratos - contratos_processados)
            minutos, segundos = divmod(int(tempo_restante), 60)
            eta = f"{minutos:02d}:{segundos:02d}"
        else:
            eta = "--:--"

        progress_label.config(text=f"{percent}% | ETA: {eta}")

    def anexar_planilha():
        nonlocal caminho_planilha
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not caminho:
            u.print_log(logs_widget, "⚠ Nenhum arquivo selecionado")
            return
        
        try:
            xls = pd.ExcelFile(caminho)
            abas = xls.sheet_names

            if len(abas) == 1:
                aba_selecionada = abas[0]
            else:
                escolha_win = tk.Toplevel(frame)
                escolha_win.title("Selecionar aba")
                escolha_win.geometry("400x250")
                escolha_win.configure(bg="#282a36")

                tk.Label(
                    escolha_win,
                    text="Selecione a aba da planilha:",
                    bg="#282a36", fg="#f8f8f2",
                    font=("Consolas", 10, "bold")
                ).pack(pady=10)

                aba_var = tk.StringVar(value=abas[0])
                combo = ttk.Combobox(escolha_win, textvariable=aba_var, values=abas, state="readonly")
                combo.pack(pady=5)

                def confirmar_aba():
                    nonlocal caminho_planilha
                    aba_escolhida = aba_var.get()
                    caminho_planilha = (caminho, aba_escolhida)
                    u.print_log(logs_widget, f"📄 Arquivo: {caminho}\n📘 Aba selecionada: {aba_escolhida}")
                    escolha_win.destroy()

                ttk.Button(escolha_win, text="Confirmar", command=confirmar_aba).pack(pady=10)
                escolha_win.grab_set()
                return  # encerra até o usuário escolher a aba
            
            caminho_planilha = (caminho, aba_selecionada)
            u.print_log(logs_widget, f"📄 Arquivo selecionado: {caminho}\n📘 Aba: {aba_selecionada}")

        except Exception as e:
            u.print_log(logs_widget, f"❌ Erro ao ler abas da planilha: {e}")


    def executar_logs_bloqueio_thread():
        nonlocal caminho_planilha
        if not caminho_planilha:
            u.print_log(logs_widget, "❌ Nenhum arquivo Excel selecionado.")
            return
        
        def target():
            nonlocal df_resultado
            try:
                # Lê a planilha e inicializa a barra de progresso
                caminho_filtrado, total_registros = logs_bloqueio.extrair_dados_planilha(
                    caminho_planilha,
                    lambda msg: u.print_log(logs_widget, msg)
                )

                if total_registros == 0:
                    u.print_log(logs_widget, "❌ Nenhum registro após filtragem.")
                    return

                iniciar_barra(total_registros)  # barra com tamanho correto

                # Processa a planilha filtrada no SAP
                df_resultado = logs_bloqueio.executar_logs_bloqueio(
                    caminho_filtrado,  # passa o arquivo filtrado
                    lambda msg: u.print_log(logs_widget, msg),
                    atualizar_progresso=lambda passo=1: frame.after(0, lambda: atualizar_barra(passo))
                )

                progress["value"] = progress["maximum"]
                atualizar_barra(0)
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