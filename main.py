#!/usr/bin/env python3

"""
Последняя версия
jan 2023
"""
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo
from tkinter.ttk import Progressbar
from tkinter import ttk
from tkinter import *
import pandas as pd
import tkinter.messagebox as mb
import time
import create_pdf as cp
from fpdf import FPDF
# Добавляем необходимые подклассы - MIME-типы
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from pathlib import Path
import os
import sys

global path_baza


def choose_file():
    global filename1
    filetypes = ("Текстовый файл", "*.xlsm"),
    filename1 = fd.askopenfilename(title="Выбрать файл",
                                   initialdir="/",
                                   filetypes=filetypes)
    return filename1


# Функция для переделки пути к файлу
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)


def create_pdf(p1, p2): #, trim, year=False, send_email=False):
    """ Функция создающая pdf-документ и записывает файл

    :param send_email:
    :param p1: Номер строки где находится АО на листе "2023"
    :param p2: Адрес электронной почты
    :param trim: номер квартала 1-4 или 0 если счёт на год
    :param year: true/false если нужен 2-й счет до конца года
    :return:

    """
    # period = [{'Servicii_de_tinere_a_registrului_Anul_2023': 49},
    #           {'Datorie': 5, 'Servicii_de_tinere_a_registrului_trim.1': 6,
    #            'Extras': 7, 'Lista': 8, 'Scrisori': 9, 'Emisie': 10,
    #            'Informatie': 11, 'Consalting': 12, 'Alte': 13,
    #            'Spre_achitare': 16},
    #           {'Datorie': 16, 'Servicii_de_tinere_a_registrului_trim.2': 17,
    #            'Extras': 18, 'Lista': 19, 'Scrisori': 20, 'Emisie': 21,
    #            'Informatie': 22, 'Consalting': 23, 'Alte': 24,
    #            'Spre_achitare': 27},
    #           {'Datorie': 27, 'Servicii_de_tinere_a_registrului_trim.3': 28,
    #            'Extras': 29, 'Lista': 30, 'Scrisori': 31, 'Emisie': 32,
    #            'Informatie': 33, 'Consalting': 34, 'Alte': 35,
    #            'Spre_achitare': 38},
    #           {'Datorie': 38, 'Servicii_de_tinere_a_registrului_trim.4': 39,
    #            'Extras': 40, 'Lista': 41, 'Scrisori': 42, 'Emisie': 43,
    #            'Informatie': 44, 'Consalting': 45, 'Alte': 46,
    #            'Spre_achitare': 49}]
    data = []
    # year = False if trim == 4 else year
    x_s = 100  # расположение печати по оси X
    y_s = -3  # расположение печати по оси Y
    # for i in range(2, len(table) - 1):
    #     print(table[1][i], table[49][i])
    #     if table.loc[p1][v]) != 'nan':
    # for k, v in period[int(4)].items():
    #     if str(table.loc[p1][v]) != 'nan':
    #         # print(type(table.loc[p1][v]))
    #         if k == "Datorie" and table.loc[p1][v] < 0:
    #             k = "Avans"
    #         data.append([k, table.loc[p1][v]])
    #         y_s += 1

    date = time.strftime("%d.%m.%Y")
    txt0 = "Anexa 13\nFormular INV-9"
    txt1 = 'EXTRAS DIN CONT'
    txt2 = 'S.A. "Grupa Financiară"'
    txt3 = 'Adresa: MD2001, mun. Chişinău, str. A.Bernardazzi, 7, of. 7'
    txt4 = 'IBAN: MD17AG000000022515490120 la B.C. ' \
           '"MOLDOVA-AGROINDBANK" S.A. suc. Tighina '
    txt5 = 'Cod bancar: AGRNMD2X864'
    txt6 = 'IDNO: 1002600054323'
    txt7 = 'Telefon: 022-27-18-45, 022-27-37-13 (contabilitatea)'

    txt6 = 'Platitor: S.A. "' + p1 + '"'
    txt_ps1 = """Если у Вашей компании есть возможность оплатить услуги за оставшийся период 2023 года, ниже представлен счёт по которому можно произвести платёж."""
    txt_ps2 = """ВНИМАНИЕ!!! Налоговая накладная на наши услуги отписывается посредством системы 'E-FACTURA' по длинному циклу. Просим своевременно подписывать её. Все вопросы и пожелания принимаются по телефону 022-27-23-13 (бухгалтерия)."""
    txt_ps3 = """IN ATENTIA CONTABILITATII!!! Factura fiscala pentru serviciile acordate SA 'Grupa Financiară' este perfectata prin intermediul 'E-FACTURA' - ciclu mare. Rugam semnarea acestei facturi in timp util. Ralatii la telefon 022-27-37-13 (contabilitatea)."""
    txt_ps4 = """Dacă doriți să achitați serviciile totale pentruanul 2023 mai jos va prezentăm contul anual."""
    txt_sign1 = 'Directoare Viorica BONDAREV'
    txt_sign2 = 'Contabila-şef Maia CIULCOVA'

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '',
                 resource_path(f"{Path('DejaVuSansCondensed.ttf')}"),
                 uni=True)
    pdf.add_font('DejaVu', 'B',
                 resource_path(f"{Path('DejaVuSans-Bold.ttf')}"),
                 uni=True)
    pdf.set_font("DejaVu", size=20, style="B")
    pdf.cell(200, 10, txt1, ln=1, align="C")
    pdf.set_font("DejaVu", size=13)
    pdf.cell(200, 6, txt2, ln=1)
    pdf.cell(200, 6, txt3, ln=1)
    pdf.cell(200, 6, txt4, ln=1)
    pdf.cell(200, 6, txt5, ln=1)
    pdf.ln()
    pdf.set_font("DejaVu", size=14, style="B")
    pdf.cell(200, 6, txt6, ln=1)
    pdf.set_font("DejaVu", size=13)

    # вывод верхней таблицы со счётом за квартал
    data_z = ['Document', 'Explicatii', 'Suma, lei']
    spacing = 1
    col_width = pdf.w / 2.22  # 4.5
    row_height = pdf.font_size * 1.2
    # Zagolovok
    pdf.set_font("DejaVu", size=13, style="B")
    for item in data_z:
        pdf.cell(col_width, row_height * spacing, txt=item, border=1,
                 align="C")
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=13)
    # print("-" * 100)
    for row in data:
        # print("ROW====", row[0], row[1])
        pdf.cell(col_width, row_height * spacing, txt=str(row[0]), border=1)
        pdf.cell(col_width, row_height * spacing, txt=f"{float(row[1]):.2f}",
                 border=1,
                 align="R")
        pdf.ln(row_height * spacing)

    pdf.ln()
    pdf.cell(200, 6, txt_sign1, ln=1)
    pdf.cell(200, 10, txt_sign2, ln=1)
    pdf.ln()
    pdf.set_font("DejaVu", size=10)
    pdf.set_text_color(220, 50, 50)
    pdf.multi_cell(0, 4, txt_ps2)
    pdf.multi_cell(0, 4, txt_ps3)
    pdf.image(resource_path('stamp.png'), x=x_s, y=67 + y_s * 6, w=38)
    pdf.image(resource_path('sign.png'), x=x_s - 25, y=80 + y_s * 6, w=13)

                 # txt="Servicii_de_tinere_a_registrului_Anul_2023", border=1)

    emitent = table.loc[p1][1]
    emitent = emitent.replace('"', "")
    emitent = emitent.replace(',', "")
    emitent = emitent.replace(' ', "_")
    emitent = emitent.replace("Ţ", "T")
    emitent = emitent.replace("Ș", "S")
    emitent = emitent.replace("Ş", "S")
    emitent = emitent.replace("Ă", "A")
    emitent = emitent.replace("Î", "I")
    emitent = emitent.replace("Â", "I")

    # Сохраняем счёт
    home = Path.home()
    if not (Path(home, 'Desktop', 'Conturi')).is_dir():
        Path(home, 'Desktop', 'Conturi').mkdir()
    path_file = f"{Path(home, 'Desktop', 'Conturi', f'Extras_din_cont_{emitent}.pdf')}"
    try:
        print(f"Saving a file to Cont_{emitent}.pdf", end="\t")
        pdf.output(path_file)
        print("Saved!")
    except UnicodeEncodeError:
        print("\nBUG! File not saved!", table.loc[p1][1], path_file)
    except FileNotFoundError:
        print("\nError! File not saved!", table.loc[p1][1], path_file)


    # Высылаем счёт если выбрано
    send_email = False
    if send_email:
        status = True
        d1 = time.strftime("%d.%m.%Y")
        t1 = time.strftime("%H:%M:%S")
        if str(p2) == 'nan':
            print("Verify email - email not found!!!")
            status = False
            tmp_status = "no email"
        # print([emitent, p2, d1, t1, status])
        # можно поменять NaN на что-то другое
        # tmp_status = "OK" if status else " "
        # report.append([emitent, str(p2), d1, t1, tmp_status])

        if status:
            email_sender = "grupa_financiara@mail.ru"
            email_password = "p0KaR11DGgTDCNycPfQs"  # нужно сгенерировать новый пароль https://id.mail.ru/security
            email_receiver = p2  # "dnedosek@gmail.com"  # p2 список адресов

            body = """
            Acest document,poate conține date personale. Dacă l-ați primit 
            din greșeală: nu aveţi dreptul de divulgare, păstrare, transmitere a acestor date. Vă rugăm să anunţaţi imediat Grupa Financiară pe e-mail, adresa juridică şi/sau la tel. 022 27-18-45.
            """
            em = MIMEMultipart()
            em['From'] = email_sender
            em['To'] = email_receiver
            em['subject'] = "Cont pentru achitarea serviciilor SA Grupa " \
                            "Financiara (SA " + table.loc[p1][1] + ")"
            em.attach(MIMEText(body, 'plain'))

            filename = "Cont_" + emitent + ".pdf"
            attachment = open(path_file, "rb")
            part = MIMEBase('application', 'pdf')  # <<<--------------------
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f"attachment; filename={filename}")
            em.attach(part)
            try:
                # print("path_file:", path_file)
                # print("email_receiver:", p2)
                with smtplib.SMTP_SSL('smtp.mail.ru', 465) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, email_receiver.split(", "),
                                  em.as_string())
                print("Email send")
                # word_editor.insert(INSERT, f"Email on {p2} sended\n")
                tmp_status = "Send email"
            except:
                print("Don't send")
                # word_editor.insert(INSERT, f"Don`t send\n")
                tmp_status = "email not sent"

        # добавляет данные в действии в конец файла report.txt
        report.append([emitent, str(p2), d1, t1, tmp_status])
        file_report = open(Path(home, 'Desktop', 'Conturi', 'report.txt'), "a")
        file_report.write(
            f"{emitent},\t\t\t{str(p2)},\t\t{d1},\t\t{t1},\t\t{tmp_status}\n")
        file_report.close()


def save_report():  # Создаём и сохраняем на диске
    # --------------------------------------------
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.add_font('DejaVu', '', f"{Path('DejaVuSansCondensed.ttf')}",
                 uni=True)
    pdf.add_font('DejaVu', 'B', f"{Path('DejaVuSans-Bold.ttf')}",
                 uni=True)

    pdf.set_font("DejaVu", size=10, style="B")

    # вывод верхней таблицы со счётом за квартал
    dat = ['Denumirea', 'email', 'Data', 'Timp', 'Send']
    spacing = 1
    # col_width = pdf.w / 2.22  # 4.5
    row_height = pdf.font_size * 1.2
    # Zagolovok
    pdf.set_font("DejaVu", size=11, style="B")
    r1 = [90, 130, 22, 18, 18]
    s = 0
    for item in dat:
        pdf.cell(r1[s], row_height * spacing, txt=item, border=1, align="C")
        s += 1
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=9)
    for row in report:
        pdf.cell(r1[0], row_height * spacing, txt=row[0], border=1)
        pdf.cell(r1[1], row_height * spacing, txt=row[1], border=1, align="R")
        pdf.cell(r1[2], row_height * spacing, txt=row[2], border=1, align="R")
        pdf.cell(r1[3], row_height * spacing, txt=row[3], border=1, align="R")
        pdf.cell(r1[4], row_height * spacing, txt=str(row[4]), border=1,
                 align="R")
        pdf.ln(row_height * spacing)

    pdf.ln()
    # print(report)
    home = Path.home()
    # p_file = Path(home, 'Desktop', 'cont2', 'Conturi', "report.pdf")
    # pdf.output(f"{Path(home, 'Desktop', 'cont2', 'Conturi', 'report.pdf')}")
    pdf.output(f"{Path(home, 'Desktop', 'Conturi', 'report.pdf')}")
    print("Report saved")
    # --------------------------------------------


def save_and_send():
    # список номеров позиций которые отмечены
    list_for_send = [i for i, x in enumerate(state) if x.get() == 1]
    # print(list_for_send)
    # Если не выбрано ни одно АО, вывести окно с ошибкой
    # Для последовательностей (строк, списков, кортежей) используйте тот факт,
    # что пустые последовательности являются ложными
    if not list_for_send:
        mb.showwarning("Ошибка", "Не выбрано ни одно АО")
        return
    start_time = time.time()
    for i in list_for_send:
        for j in range(1, len(table_email)):
            if table_email.loc[j][1] == table[1][i]:
                create_pdf(i, table_email.loc[j][8],
                           kvartal.get(),
                           var_endY.get(),
                           var_cb2.get())
    save_report()
    end_time = time.time() - start_time
    m = end_time // 60
    s = end_time % 60
    showinfo(title="Информация",
             message=f"Обработано АО: {set_checkbox()}\n"
                     f"Время обработки: {int(m)} мин. {int(s)} сек.")


def conts_for_all():
    for i in range(2, len(table) - 1):
        trim_column = [49, 16, 27, 38, 49]  # Номера столбцов с задолженностями
        if str(table.loc[i][2]) != "STOP" and str(
                table.loc[i][2]) != "EXPIRAT":
            if str(table.loc[i][trim_column[49]]) != 'nan' and int(
                    table.loc[i][trim_column[kvartal.get()]]) > 0:
                state[i].set(True)

    save_and_send()


def sa_choice():
    """ Установка фложков АО у которых значение в выбранном квартале > 0 """
    for sa in range(2, len(table) - 1):
        trim_column = [49, 16, 27, 38, 49]  # Номера столбцов с задолженностями
        if str(table.loc[sa][2]) != "STOP" and str(
                table.loc[sa][2]) != "EXPIRAT":
            if str(table.loc[sa][trim_column[49]]) != 'nan' and int(
                    table.loc[sa][trim_column[kvartal.get()]]) > 0:
                state[sa].set(True)
    # set_checkbox()



# ----- Start programm -----
print("start")
# global path_baza
initial_dir = os.getcwd()
file_path = os.path.join(initial_dir, 'БАЗА.xlsm')
if os.path.exists(file_path):
    path_baza = resource_path(Path(file_path))
else:
    filetypes = ("Excel", "*.xlsm"),
    filename = fd.askopenfilename(title="Выбрать файл",
                                  initialdir=initial_dir,
                                  filetypes=filetypes)
    if filename == "":
        mb.showinfo("Не выбран файл")
    path_baza = resource_path(Path(filename))

report = []
# global table, table_email
table = pd.DataFrame(
    pd.read_excel(path_baza, sheet_name='2022', header=None))
table_email = pd.DataFrame(
    pd.read_excel(path_baza, sheet_name='data', header=None))
for i in range(2, len(table)-1):
    print(table[1][i], table[49][i])
print(table)
print(table_email)
# for i in range(2, 6):
#     cp.create_pdf(table[1][i], "dnedosek@gmail.com")

cp.create_pdf("VIV", "dnedosek@gmail.com")