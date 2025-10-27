from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import xmltodict
import style
import utils as u
import threading
import time

interromper = False  # variável global de controle


def criar_frame_conversor_parquet(parent, btn_voltar=None):
    frame = ttk.Frame(parent, padding=10)
    conteudo_frame = ttk.Frame(frame)
    conteudo_frame.pack(fill="both", expand=True)

    arquivos_selecionados = []

    # ---------------- LOGS ----------------
    def log(msg):
        logs_widget.after(0, lambda: u.print_log(logs_widget, msg))

    # ---------------- SELECIONAR ARQUIVO ----------------
    def selecionar_arquivo():
        file = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[("Arquivos suportados", "*.xml *.csv *.json *.xlsx")]
        )
        if file:
            arquivos_selecionados.clear()
            arquivos_selecionados.append(file)
            log(f"Arquivo selecionado: {file}")
            label_progresso.config(text="Arquivo pronto para conversão.")

    ttk.Button(conteudo_frame, text="Selecionar arquivo", command=selecionar_arquivo).pack(pady=5)

    # ---------------- ELEMENTO XML ----------------
    ttk.Label(conteudo_frame, text="Elemento raiz do XML (ex: livraria/livro):").pack(pady=5)
    entrada_elemento_xml = ttk.Entry(conteudo_frame, width=50)
    entrada_elemento_xml.pack(pady=5)

    # ---------------- LABEL DE PROGRESSO ----------------
    label_progresso = ttk.Label(conteudo_frame, text="", font=("Segoe UI", 10, "italic"))
    label_progresso.pack(pady=5)

    # ---------------- LOGS ----------------
    logs_frame = ttk.Frame(conteudo_frame)
    logs_frame.pack(fill="both", expand=True, pady=5)
    logs_widget = style.criar_logs_widget(logs_frame, width=90, height=10)
    logs_widget.pack(fill="both", expand=True)

    # ---------------- CONVERSÃO ----------------
    def converter_para_parquet():
        global interromper
        interromper = False

        if not arquivos_selecionados:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
            return

        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".parquet",
            filetypes=[("Parquet", "*.parquet")]
        )
        if not caminho_saida:
            return

        elemento_xml = entrada_elemento_xml.get().strip()
        arquivo = arquivos_selecionados[0]
        ext = os.path.splitext(arquivo)[1].lower()

        log(f"Iniciando conversão: {arquivo}")
        logs_widget.after(0, lambda: label_progresso.config(text="Carregando dados..."))

        try:
            t0 = time.time()

            # função para atualizar o label de progresso
            def atualizar_progresso(linhas_lidas, total_estimado=None):
                if total_estimado:
                    perc = (linhas_lidas / total_estimado) * 100
                    label_text = f"Lidas {linhas_lidas:,}/{total_estimado:,} linhas ({perc:.1f}%)"
                else:
                    label_text = f"Lidas {linhas_lidas:,} linhas..."
                label_progresso.after(0, lambda: label_progresso.config(text=label_text))

            # ------------------- LEITURA DO ARQUIVO -------------------
            df_total = None
            linhas_lidas = 0

            if ext == ".csv":
                # leitura em chunks (não trava)
                for chunk in pd.read_csv(arquivo, chunksize=5000):
                    if interromper:
                        log("Conversão interrompida pelo usuário.")
                        label_progresso.after(0, lambda: label_progresso.config(text="Conversão interrompida."))
                        return
                    linhas_lidas += len(chunk)
                    atualizar_progresso(linhas_lidas)
                    if df_total is None:
                        df_total = chunk
                    else:
                        df_total = pd.concat([df_total, chunk], ignore_index=True)

            elif ext == ".json":
                df_total = pd.read_json(arquivo)
                atualizar_progresso(len(df_total), len(df_total))

            elif ext in [".xls", ".xlsx"]:
                # lê todas as abas de uma vez
                sheets = pd.read_excel(arquivo, sheet_name=None)

                # total de linhas para progresso
                total_estimado = sum(len(df) for df in sheets.values())
                dfs = []
                linhas_lidas = 0

                # processa cada aba
                for nome, df in sheets.items():
                    if interromper:
                        log("Conversão interrompida pelo usuário.")
                        label_progresso.after(0, lambda: label_progresso.config(text="Conversão interrompida."))
                        return

                    linhas_lidas += len(df)
                    atualizar_progresso(linhas_lidas, total_estimado)  # atualização “rápida” de progresso
                    dfs.append(df)

                df_total = pd.concat(dfs, ignore_index=True)

                # tenta salvar e corrige colunas problemáticas dinamicamente
                while True:
                    try:
                        df_total.to_parquet(caminho_saida, engine="pyarrow", index=False)
                        break  # salvamento concluído
                    except Exception as e:
                        msg = str(e)
                        if "Conversion failed for column" in msg:
                            import re
                            col = re.search(r"column (\w+)", msg)
                            if col:
                                col_name = col.group(1)
                                log(f"Forçando coluna {col_name} como texto (str) devido a erro de conversão...")
                                df_total[col_name] = df_total[col_name].astype(str)
                            else:
                                raise e
                        else:
                            raise e

            elif ext == ".xml":
                if not elemento_xml:
                    messagebox.showerror("Erro", "Para XML é necessário informar o elemento raiz.")
                    return
                with open(arquivo, 'r', encoding='utf-8') as f:
                    xml_dict = xmltodict.parse(f.read())
                data_list = xml_dict
                for tag in elemento_xml.split('/'):
                    if isinstance(data_list, dict):
                        data_list = data_list.get(tag, [])
                    else:
                        data_list = []
                if not isinstance(data_list, list):
                    data_list = [data_list]
                df_total = pd.DataFrame(data_list)
                atualizar_progresso(len(df_total), len(df_total))

            else:
                messagebox.showerror("Erro", f"Tipo de arquivo não suportado: {ext}")
                return

            # ------------------- SALVAMENTO -------------------
            log("Salvando arquivo Parquet...")
            label_progresso.after(0, lambda: label_progresso.config(text="Salvando Parquet..."))
            df_total.to_parquet(caminho_saida, engine="pyarrow", index=False)

            tempo_total = time.time() - t0
            log(f"Conversão concluída em {tempo_total:.1f}s: {caminho_saida}")
            label_progresso.after(0, lambda: label_progresso.config(
                text=f"Conversão concluída em {tempo_total:.1f}s — {len(df_total):,} linhas.")
            )
            messagebox.showinfo("Sucesso", f"Parquet gerado com sucesso:\n{caminho_saida}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
            log(f"Erro: {str(e)}")

    # ---------------- THREADING E BOTÕES ----------------
    botoes_frame = ttk.Frame(conteudo_frame)
    botoes_frame.pack(pady=10)

    btn_converter = ttk.Button(botoes_frame, text="Converter para Parquet")
    btn_converter.pack(side="left", padx=5)

    btn_interromper = ttk.Button(botoes_frame, text="Interromper conversão", state='disabled')
    btn_interromper.pack(side="left", padx=5)

    def thread_converter():
        global interromper
        interromper = False
        btn_converter.config(state="disabled")
        btn_interromper.config(state="normal")

        def run():
            try:
                converter_para_parquet()
            finally:
                btn_converter.config(state="normal")
                btn_interromper.config(state="disabled")

        threading.Thread(target=run).start()

    def interromper_conversao():
        global interromper
        interromper = True
        label_progresso.config(text="Solicitando interrupção...")
        log("Solicitando interrupção...")

    btn_converter.config(command=thread_converter)
    btn_interromper.config(command=interromper_conversao)

    style.aplicar_estilo(frame)
    return frame, logs_widget, None
