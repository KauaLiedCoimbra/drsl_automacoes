import tkinter as tk
from tkinter import ttk, messagebox
import threading
import style as s
from .auto_enter import auto_enter
from PIL import Image, ImageTk
import os


def criar_frame_auto_enter(parent, btn_voltar=None):
    frame = ttk.Frame(parent, padding=10, style="Custom.TFrame")

    intervalo_var = tk.StringVar(value="0.3")
    duracao_var = tk.StringVar(value="10")
    key_var = tk.StringVar(value="ENTER")
    turbo_var = tk.BooleanVar(value=False)
    interromper = tk.BooleanVar(value=False)

    # Estilos
    style = ttk.Style()
    style.configure("Custom.TFrame", background=s.DRACULA_BG)
    style.configure("Custom.TLabel", background=s.DRACULA_BG, foreground=s.DRACULA_FG)
    style.configure("Custom.TCheckbutton", background=s.DRACULA_BG, foreground=s.DRACULA_FG)

    # LOGS
    logs_widget = tk.Text(frame, height=12, width=70, state="disabled",
                          bg=s.DRACULA_LOGS_WIDGET, fg=s.DRACULA_FG, relief="flat",
                          wrap="word", insertbackground=s.DRACULA_FG)
    logs_widget.pack(fill="both", expand=True, padx=5, pady=(0, 10))

    # Container para esquerda (config) + direita (gif)
    container = ttk.Frame(frame, style="Custom.TFrame")
    container.pack(fill="x", pady=10)

    # ---- ESQUERDA: CONFIG ----
    config = ttk.Frame(container, style="Custom.TFrame")
    config.pack(side="left", anchor="n", padx=10)

    ttk.Label(config, text="Tecla:", style="Custom.TLabel").grid(row=0, column=0, sticky="e")
    ttk.Combobox(config, textvariable=key_var,
                 values=["ENTER", "F3", "F8", "Confirma (✔)", "Page Down", "Page Up"],
                 width=14, state="readonly").grid(row=0, column=1, padx=5)

    ttk.Label(config, text="Intervalo (s):", style="Custom.TLabel").grid(row=1, column=0, sticky="e")
    tk.Entry(config, textvariable=intervalo_var, width=8,
             bg=s.DRACULA_WIDGET_BG, fg=s.DRACULA_FG,
             insertbackground=s.DRACULA_FG).grid(row=1, column=1, padx=5)

    ttk.Label(config, text="Duração (s):", style="Custom.TLabel").grid(row=2, column=0, sticky="e")
    tk.Entry(config, textvariable=duracao_var, width=8,
             bg=s.DRACULA_WIDGET_BG, fg=s.DRACULA_FG,
             insertbackground=s.DRACULA_FG).grid(row=2, column=1, padx=5)

    ttk.Checkbutton(config, text="Modo Turbo (ignorar intervalo)", variable=turbo_var,
                    style="Custom.TCheckbutton").grid(row=3, column=0, columnspan=2, pady=5)

    # ---- DIREITA: GIF ----
    gif_frame = ttk.Frame(container, style="Custom.TFrame")

    gif_path = os.path.join(os.path.dirname(__file__), "auto_enter_gif.gif")
    if not os.path.exists(gif_path):
        logs_widget.config(state="normal")
        logs_widget.insert("end", "[ERRO] GIF não encontrado: auto_enter_gif.gif\n")
        logs_widget.config(state="disabled")
        gif_frame = None
    else:
        gif_image = Image.open(gif_path)

        scale = 0.7

        frames = []
        try:
            while True:
                frame_img = gif_image.copy()

                # Redimensiona
                new_width = int(frame_img.width * scale)
                new_height = int(frame_img.height * scale)
                frame_img = frame_img.resize((new_width, new_height), Image.LANCZOS)

                frames.append(ImageTk.PhotoImage(frame_img))
                gif_image.seek(len(frames))  # vai para o próximo frame
        except EOFError:
            pass

        gif_label = tk.Label(gif_frame, bg=s.DRACULA_BG)
        gif_label.pack()

        # começa invisível
        gif_frame_is_visible = False
        animacao_id = None

        def mostrar_gif():
            nonlocal gif_frame_is_visible
            if not gif_frame_is_visible:
                gif_frame.pack(side="right", padx=30, pady=5)
                gif_frame_is_visible = True
                animar(0)  # ← sempre reinicia do frame 0

        def esconder_gif():
            nonlocal gif_frame_is_visible, animacao_id
            if gif_frame_is_visible:
                gif_frame.pack_forget()
                gif_frame_is_visible = False
                if animacao_id is not None:
                    frame.after_cancel(animacao_id)
                    animacao_id = None

        def animar(ind=0):
            nonlocal animacao_id
            if gif_frame_is_visible:
                gif_label.configure(image=frames[ind])
                animacao_id = frame.after(60, animar, (ind + 1) % len(frames))

    # ---- BOTÕES ----
    btns = ttk.Frame(frame, style="Custom.TFrame")
    btns.pack(pady=10)

    def executar():
        try:
            intervalo = float(intervalo_var.get().replace(",", "."))
            duracao = float(duracao_var.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Erro", "Digite valores numéricos.")
            return

        interromper.set(False)

        # mostrar GIF
        mostrar_gif()

        def run():
            auto_enter(key_var.get(), intervalo, duracao, turbo_var.get(), logs_widget, interromper)
            gif_frame.pack_forget()

        threading.Thread(target=run, daemon=True).start()

    tk.Button(btns, text="▶ Executar", bg=s.DRACULA_BUTTON_BG, fg=s.DRACULA_FG,
              activebackground=s.DRACULA_BUTTON_ACTIVE, command=executar).pack(side="left", padx=5)

    def parar():
        interromper.set(True)
        esconder_gif()

    tk.Button(btns, text="⛔ Interromper", bg="#ff5555", fg=s.DRACULA_FG,
              activebackground="#ff6e6e", command=parar).pack(side="left", padx=5)

    # VOLTAR
    if btn_voltar:
        btn_voltar.pack(pady=10)

    return frame, logs_widget, interromper
