import time
from fpdf import FPDF
import pandas as pd
import tkinter.messagebox as mb
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo
# Добавляем необходимые подклассы - MIME-типы
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from pathlib import Path
import os
import sys


def create_pdf(sa, value, email):
    """
    Функция создающая pdf-документ и записывает файл
    """
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', f"{Path('DejaVuSansCondensed.ttf')}", uni=True)
    pdf.add_font('DejaVu', 'B', f"{Path('DejaVuSans-Bold.ttf')}", uni=True)
    pdf.set_font("DejaVu")
    pdf.cell(190, 6, txt='Anexa 13', ln=1, align="R")
    pdf.cell(190, 6, txt='Formular INV-9', ln=1, align="R")
    pdf.ln()
    pdf.cell(95, 6, txt='S.A. "Grupa Financiară"')
    pdf.cell(95, 6, txt=f'S.A. "{sa}"', ln=1, align="R")
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
    if value < 0:
        current_account = f"înregistrează creanţe în sumă de " \
                          f"{float(abs(value)):.2f} lei."
    elif value > 0:
        current_account = f"înregistrează datorii în sumă de " \
                          f"{float(value):.2f} lei."
    else:
        current_account = f"înregistrează datorii/creanţe în sumă de 0.00 lei."
    pdf.multi_cell(190, 6, txt="Conform Ordinului privind efectuarea "
                               "inventarierii nr.22 din 20.12.2022, Vă "
                               "comunicăm, că în contabilitatea S.A. "
                               "'Grupa Financiară', la data de 31 decembrie "
                               "2022, entitate Dvs. entitatea Dvs. " +
                               current_account)
    # pdf.set_font("DejaVu", size=14, style="B")
    # pdf.cell(200, 6, txt6, ln=1)
    # pdf.set_font("DejaVu", size=13)

    # вывод верхней таблицы со счётом за квартал
    data_z = ['Document', 'Explicatii', 'Suma, lei']
    spacing = 1
    col_width = pdf.w / 3.33
    row_height = pdf.font_size * 2
    # Zagolovok
    pdf.set_font("DejaVu", size=13, style="B")
    for item in data_z:
        pdf.cell(col_width, row_height * spacing, txt=item, border=1,
                 align="C")
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=13)

    pdf.cell(col_width, row_height * spacing, txt='Contract', border=1)
    pdf.cell(col_width, row_height * spacing, txt='Servicii de tinere '
                                                  'registrului', border=1)
    pdf.cell(col_width, row_height * spacing, txt=f"{float(value):.2f}",
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
    pdf.image('sign.png', x=72, y=198, w=13)

    # Сохраняем счёт
    home = Path.home()
    if not (Path(home, 'Desktop', 'Conturi')).is_dir():
        Path(home, 'Desktop', 'Conturi').mkdir()
    path_file = f"{Path(home, 'Desktop', 'Conturi', f'Extras_din_cont_{sa}.pdf')}"
    try:
        print(f"Saving a file to Extras_din_cont_{sa}.pdf", end="\t")
        pdf.output(path_file)
        print("Saved!")
    except UnicodeEncodeError:
        print("\nBUG! File not saved!", sa, path_file)
    except FileNotFoundError:
        print("\nError! File not saved!", sa, path_file)

    # Высылаем счёт если выбрано
    if send_email:
        status = True
        d1 = time.strftime("%d.%m.%Y")
        t1 = time.strftime("%H:%M:%S")
        if str(email) == 'nan':
            print("Verify email - email not found!!!")
            status = False
            tmp_status = "no email"
        # можно поменять NaN на что-то другое

        if status:
            email_sender = "grupa_financiara@mail.ru"
            # нужно сгенерировать новый пароль https://id.mail.ru/security
            email_password = "p0KaR11DGgTDCNycPfQs"
            email_receiver = 'dnedosek@gmail.com'   # email
            # email  # "dnedosek@gmail.com"  # email список адресов

            body = "Acest document, poate conține date personale. Dacă l-ați " \
                   "primit din greșeală: nu aveţi dreptul de divulgare, " \
                   "păstrare, transmitere a acestor date. Vă rugăm să " \
                   "anunţaţi imediat Grupa Financiară pe e-mail, adresa " \
                   "juridică şi/sau la tel. 022-27-18-45."
            em = MIMEMultipart()
            em['From'] = email_sender
            em['To'] = email_receiver
            em['subject'] = "Extras din cont SA " + sa
            em.attach(MIMEText(body, 'plain'))

            filename = "Extras_din_cont_" + sa + ".pdf"
            attachment = open(path_file, "rb")
            part = MIMEBase('application', 'pdf')  # <<<--------------------
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            f"attachment; filename={filename}")
            em.attach(part)
            try:
                with smtplib.SMTP_SSL('smtp.mail.ru', 465) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, email_receiver.split(", "),
                                  em.as_string())
                print("Email send")
                tmp_status = "Send email"
            except:
                print("Don't send")
                tmp_status = "email not sent"

        # добавляет данные в действии в конец файла report.txt
        report.append([sa, str(email), d1, t1, tmp_status])
        file_report = open(Path(home, 'Desktop', 'Conturi', 'report.txt'), "a")
        file_report.write(
            f"{sa},\t\t\t{str(email)},\t\t{d1},\t\t{t1},\t\t{tmp_status}\n")
        file_report.close()


# Функция для переделки пути к файлу
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.getcwd()
    return os.path.join(base_path, relative_path)
# --------------------------------------------


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
    print(report)
    home = Path.home()
    pdf.output(f"{Path(home, 'Desktop', 'Conturi', 'report.pdf')}")
    print("Report saved")
    # --------------------------------------------


# ----- Start programm -----
start_time = time.time()
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

send_email = True   # False  # !!!!!!!!!! для отсылки писем установить True
counter = 0
for d in data:
    if counter > 10:
        break
    print(d)
    create_pdf(sa=d[0], value=d[1], email=d[2])
    counter += 1

save_report()
end_time = time.time() - start_time
m = end_time // 60
s = end_time % 60
showinfo(title="Информация",
         message=f"Обработано АО: {len(data)}\n"
                 f"Время обработки: {int(m)} мин. {int(s)} сек.")