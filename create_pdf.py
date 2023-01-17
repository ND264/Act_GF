import time
from fpdf import FPDF
import pandas as pd
import tkinter.messagebox as mb
import tkinter.filedialog as fd
# Добавляем необходимые подклассы - MIME-типы
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from pathlib import Path
import os
import sys

def create_pdf(p1="VIV", p2="dnedosek@gmail.com"): #, trim, year=False, send_email=False):
    """ Функция создающая pdf-документ и записывает файл

    :param send_email:
    :param p1: Номер строки где находится АО на листе "2023"
    :param p2: Адрес электронной почты
    :param trim: номер квартала 1-4 или 0 если счёт на год
    :param year: true/false если нужен 2-й счет до конца года
    :return:

    """
    period = [{'Servicii_de_tinere_a_registrului_Anul_2023': 49},
              {'Datorie': 5, 'Servicii_de_tinere_a_registrului_trim.1': 6,
               'Extras': 7, 'Lista': 8, 'Scrisori': 9, 'Emisie': 10,
               'Informatie': 11, 'Consalting': 12, 'Alte': 13,
               'Spre_achitare': 16},
              {'Datorie': 16, 'Servicii_de_tinere_a_registrului_trim.2': 17,
               'Extras': 18, 'Lista': 19, 'Scrisori': 20, 'Emisie': 21,
               'Informatie': 22, 'Consalting': 23, 'Alte': 24,
               'Spre_achitare': 27},
              {'Datorie': 27, 'Servicii_de_tinere_a_registrului_trim.3': 28,
               'Extras': 29, 'Lista': 30, 'Scrisori': 31, 'Emisie': 32,
               'Informatie': 33, 'Consalting': 34, 'Alte': 35,
               'Spre_achitare': 38},
              {'Datorie': 38, 'Servicii_de_tinere_a_registrului_trim.4': 39,
               'Extras': 40, 'Lista': 41, 'Scrisori': 42, 'Emisie': 43,
               'Informatie': 44, 'Consalting': 45, 'Alte': 46,
               'Spre_achitare': 49}]
    # data = []
    # year = False if trim == 4 else year
    # x_s = 100  # расположение печати по оси X
    # y_s = -3  # расположение печати по оси Y
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
    p1 = "VIV"
    date = time.strftime("%d.%m.%Y")

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', f"{Path('DejaVuSansCondensed.ttf')}", uni=True)
    pdf.add_font('DejaVu', 'B', f"{Path('DejaVuSans-Bold.ttf')}", uni=True)
    pdf.set_font("DejaVu")
    pdf.cell(190, 6, txt='Anexa 13', ln=1, align="R")
    pdf.cell(190, 6, txt='Formular INV-9', ln=1, align="R")
    pdf.ln()
    pdf.cell(95, 6, txt='S.A. "Grupa Financiară"')
    pdf.cell(95, 6, txt=f'S.A. "{p1}"', ln=1, align="R")
    pdf.cell(95, 6, txt='Adresa: MD2001, mun. Chişinău,')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 6, txt='str. A.Bernardazzi, 7, of. 7')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 6, txt='IBAN: MD17AG000000022515490120')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 6, txt='"MOLDOVA-AGROINDBANK" S.A. suc. Tighina')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 6, txt='Cod bancar: AGRNMD2X864')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 6, txt='IDNO: 1002600054323')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 6, txt='Telefon: 022-27-18-45, 022-27-37-13 (contabilitatea)')
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.ln()
    pdf.set_font("DejaVu", size=20, style="B")
    pdf.cell(190, 10, txt='EXTRAS DIN CONT', ln=1, align="C")
    pdf.set_font("DejaVu", size=13)
    # pdf.ln()
    pdf.multi_cell(190, 6, txt="""Conform Ordinului privind efectuarea inventarierii nr.22 din 20.12.2022, Vă comunicăm, că în contabilitatea S.A. "Grupa Financiară", la data de 31 decembrie 2022, entitate Dvs. entitatea Dvs. înregistrează creanţe în sumă de 0.00 lei și/sau datorii în sumă de 123.00 lei.""")
    # pdf.set_font("DejaVu", size=14, style="B")
    # pdf.cell(200, 6, txt6, ln=1)
    # pdf.set_font("DejaVu", size=13)

    # вывод верхней таблицы со счётом за квартал
    data_z = ['Document', 'Explicatii', 'Suma, lei']
    spacing = 1
    col_width = pdf.w / 3.33  # 4.5
    row_height = pdf.font_size * 2
    # Zagolovok
    pdf.set_font("DejaVu", size=13, style="B")
    for item in data_z:
        pdf.cell(col_width, row_height * spacing, txt=item, border=1,
                 align="C")
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=13)
    # print("-" * 100)
    # for row in data:
    #     # print("ROW====", row[0], row[1])
    #     pdf.cell(col_width, row_height * spacing, txt=str(row[0]), border=1)
    #     pdf.cell(col_width, row_height * spacing, txt=f"{float(row[1]):.2f}",
    #              border=1,
    #              align="R")
    #     pdf.ln(row_height * spacing)
    row = 123.44
    pdf.cell(col_width, row_height * spacing, txt='Contract', border=1)
    pdf.cell(col_width, row_height * spacing, txt='Servicii de tinere '
                                                  'registrului', border=1)
    pdf.cell(col_width, row_height * spacing, txt=f"{float(row):.2f}",
             border=1, align="C")
    pdf.ln(14)
    pdf.multi_cell(0, 6, txt="În termen de 5 zile de la primirea prezentului "
                             "extras, urmează să ne restituiţi Extrasul de "
                             "cont cu confirmarea sumei creanţei şi/sau "
                             "datoriei, iar în caz de nerecunoaştere a unei "
                             "sume total sau parţial, să anexaţi Nota "
                             "explicativă cuprinzînd obiecţiile Dvs.")
    pdf.multi_cell(0, 6, txt="Extrasul din cont foate fi restituit prin email "
                             "grupa_financiara@mail.ru sau prin poșta.")
    pdf.multi_cell(0, 6, txt="În caz de neprezentare în termen de 5 zile a "
                             "extrasului de cont, vom considera sumele expuse "
                             "ca fiind acceptate de întraprinderea Dvs., fapt "
                             "care nu vă admitere efectuarea oricărui "
                             "recalcul, iar orice dezacord sau pretenție nu "
                             "vor fi luare în considerație.")
    pdf.multi_cell(0, 6, txt="В случае не предоставления акта в течение 5 "
                             "дней, сальдо будет принято как подтверждение и "
                             "в дальнейшем претензии на изменение суммы не "
                             "будет браться во внимание.")
    pdf.ln()
    pdf.cell(95, 6, txt='Directoare Viorica BONDAREV', align="L")
    pdf.cell(95, 6, txt='_________________________________', ln=1, align="R")
    pdf.cell(95, 10, txt='Contabila-şef Maia CIULCOVA', align="L")
    pdf.cell(95, 10, txt='_________________________________', ln=1, align="R")
    pdf.ln()
    pdf.cell(95, 6, txt='L.Ș.', align="L")
    pdf.cell(95, 6, txt='L.Ș.', ln=1, align="R")
    pdf.image('stamp.png', x=15, y=205, w=38)
    pdf.image('sign.png', x=72, y=204, w=13)

    emitent = p1 #table.loc[p1][1]
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
        print(f"Saving a file to Extras_din_cont_{emitent}.pdf", end="\t")
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

# Функция для переделки пути к файлу
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)

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
table = pd.DataFrame(pd.read_excel(path_baza, sheet_name='2022', header=None))
table_email = pd.DataFrame(pd.read_excel(path_baza, sheet_name='data', header=None))
# for i in range(2, len(table)-1):
#     print(table[1][i], table[49][i])
# print(table)
print(table_email[8])
# for i in range(2, 6):
#     cp.create_pdf(table[1][i], "dnedosek@gmail.com")

data = []
for i in range(2, len(table) - 1):
    if str(table.loc[i][2]) != "STOP" and str(table.loc[i][2]) != "EXPIRAT":
        if str(table.loc[i][49]) != 'nan':
            for j in range(1, len(table_email)):
                if table_email.loc[j][1] == table.loc[i][1]:
                    data.append([table.loc[i][1], table.loc[i][49], table_email.loc[j][8]])

for d in data:
    print(d)
# create_pdf()