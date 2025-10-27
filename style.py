from tkinter import ttk, scrolledtext

# Cores Dracula
DRACULA_BG = "#282A36"
DRACULA_WIDGET_BG = "#44475a"
DRACULA_LOGS_WIDGET = "#1e1f29"
DRACULA_FG = "#F8F8F2"
DRACULA_TITLE = "#ff79c6"
DRACULA_BUTTON_BG = "#6272a4"
DRACULA_BUTTON_ACTIVE = "#8be9fd"

def aplicar_estilo(root):
    style = ttk.Style(root)
    style.theme_use('clam')
    
    style.configure('TFrame', background=DRACULA_BG)
    style.configure('Inner.TFrame', background=DRACULA_WIDGET_BG)
    
    style.configure('TLabel', background=DRACULA_BG, foreground=DRACULA_FG, font=('Consolas', 12))
    style.configure('Title.TLabel', background=DRACULA_BG, foreground=DRACULA_TITLE, font=('Consolas', 18, 'bold'))
    
    style.configure('TButton', background=DRACULA_BUTTON_BG, foreground=DRACULA_FG, font=('Consolas', 11, 'bold'))
    style.map('TButton',
              background=[('active', DRACULA_BUTTON_ACTIVE)],
              foreground=[('disabled', 'gray')])
    
def criar_logs_widget(parent, width=90, height=15, padx=5, pady=5, font=("Consolas", 10)):
    logs_widget = scrolledtext.ScrolledText(
        parent,
        width=width,
        height=height,
        font=font,
        fg=DRACULA_FG,
        bg=DRACULA_LOGS_WIDGET,
        relief="flat",
        borderwidth=5,
        padx=padx,
        pady=pady,
    )
    logs_widget.pack(fill="both", expand=True)
    logs_widget.config(state="disabled")
    return logs_widget