import threading
import PySimpleGUI as sg
import bot

def lanzar_bot():
    try:
        bot.main()
    except Exception as e:
        print(f"[ERROR] {e}")

layout = [
    [sg.Text("Bot de Seguimiento UNAM", font=("Segoe UI", 12, "bold"))],
    [sg.Multiline(size=(100, 26), key="-LOG-", autoscroll=True,
                  write_only=True, reroute_stdout=True, reroute_stderr=True)],
    [sg.Button("Iniciar", key="-RUN-"), sg.Button("Salir")]
]
window = sg.Window("Bot UNAM", layout, finalize=True)

while True:
    ev, _ = window.read()
    if ev in (sg.WINDOW_CLOSED, "Salir"):
        break
    if ev == "-RUN-":
        print("Iniciando bot…")
        threading.Thread(target=lanzar_bot, daemon=True).start()

window.close()
