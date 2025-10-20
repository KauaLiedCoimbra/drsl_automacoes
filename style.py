from tkinter import ttk

# Cores Dracula
DRACULA_BG = "#282A36"
DRACULA_WIDGET_BG = "#44475a"
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