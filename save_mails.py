import win32com.client
import pandas as pd
import datetime

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")


def selecciono_bandejas():
    bandejas = []
    for bandeja in range(5, 6+1):  # 5 es bandeja de salida y 6 es bandeja de entrada
        try:
            bandejas.append(mapi.GetDefaultFolder(bandeja))
        except:
            pass
    return bandejas


def recorre_mensajes(message, bodys, asunto, remitente, mail_remitente, horario):
    try:
        asunto.append(message.subject)
    except:
        asunto.append("sin asunto")
    try:
        bodys.append(message.body)
    except:
        bodys.append("sin cuerpo")
    try:
        remitente.append(message.SenderName)
    except:
        remitente.append("sin remitente")
    try:
        mail_remitente.append(message.SenderEmailAddress)
    except:
        mail_remitente.append("sin remitente")
    try:
        horario.append(message.ReceivedTime.isoformat())
    except:
        horario.append("-")
    return bodys, asunto, remitente, mail_remitente, horario


def recorre_bandejas(bodys, asunto, remitente, mail_remitente, horario, carpeta):
    for bandeja in bandejas:
        messages = bandeja.Items
        for message in list(messages):
            carpeta.append(bandeja.Name)
            bodys, asunto, remitente, mail_remitente, horario = recorre_mensajes(
                message, bodys, asunto, remitente, mail_remitente, horario)
        for folder in bandeja.Folders:
            messages = folder.Items
            for message in list(messages):
                bodys, asunto, remitente, mail_remitente, horario = recorre_mensajes(
                    message, bodys, asunto, remitente, mail_remitente, horario)
                carpeta.append(folder.Name)
    return bodys, asunto, remitente, mail_remitente, horario, carpeta


if __name__ == "__main__":
    bodys, asunto, remitente, mail_remitente, horario, carpeta = [], [], [], [], [], []
    bandejas = selecciono_bandejas()
    bodys, asunto, remitente, mail_remitente, horario, carpeta = recorre_bandejas(
        bodys, asunto, remitente, mail_remitente, horario, carpeta)
    mails = pd.DataFrame(list(zip(carpeta, horario, remitente, mail_remitente, asunto, bodys)),
                         columns=['carpeta', 'horario', 'remitente', 'mail_remitente', 'asunto', 'bodys'])
    fecha = datetime.date.today().isoformat()
    mails.to_csv(f"mails-{fecha}.csv", index=False)
