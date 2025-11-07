import time
import pythoncom
import utils as u


VKEY_MAP = {
    "ENTER": 0,
    "F3": 3,
    "F8": 8,
    "Confirma (✔)": 11,
    "Page Down": 82,
    "Page Up": 81,
}


def auto_enter(key_name, intervalo, duracao, turbo, logs_widget, interromper_var):
    pythoncom.CoInitialize()

    try:
        session = u.conectar_sap()
    except Exception as e:
        u.print_log(logs_widget, f"❌ Erro ao conectar ao SAP: {e}")
        return

    vkey = VKEY_MAP.get(key_name, 0)

    u.print_log(logs_widget, f"▶ Iniciando auto {key_name} | Intervalo: {intervalo}s | Duração: {duracao}s | Turbo={turbo}")

    inicio = time.time()
    contador = 0

    while time.time() - inicio < duracao:
        if interromper_var.get():
            u.print_log(logs_widget, f"⛔ Interrompido pelo usuário após {contador} ações.")
            return

        # Espera o SAP processar
        while session.Busy:
            time.sleep(0.03)

        try:
            session.findById("wnd[0]").sendVKey(vkey)
        except Exception:
            u.print_log(logs_widget, "⚠️ Não foi possível enviar comando (SAP pode ter perdido foco).")
            time.sleep(intervalo)
            continue

        contador += 1

        if not turbo:
            time.sleep(intervalo)

    u.print_log(logs_widget, f"✅ Finalizado. Total enviado: {contador}.")
