""" Code for operating PDF and Excel files."""

__author__ = "Josip Sivrić"
__version__ = "1.1.3.0"
__email__ = "josipsivric@gmail.com"
__status__ = "Production"

import pdfplumber
import xlwings as xw
from itertools import chain


def open_first_pdf(selected_pdf):
    """ Function for specific file. Only works properly if format of the file is not changed.

    :param selected_pdf: path to PDF to be parsed
    :return:
    """
    pdf = pdfplumber.open(selected_pdf)
    table = []

    for page in pdf.pages:
        extract = page.extract_table()
        if extract is not None:
            table.extend(extract)

    new_table = table[2:]
    formated_table = [[""] * 18 for _ in range(len(new_table))]
    positions = [0, 1, 13, 8, 7, 10, 11, 14, 12]
    deduplicated_table = []
    for i in range(len(new_table)):
        new_table[i][3] = new_table[i][3].replace(',', '.')
        new_table[i][4] = new_table[i][4].replace(',', '.')

    for i in range(len(new_table)):
        for index, pos in enumerate(positions):
            formated_table[i][pos] = new_table[i][index]

    for i in range(len(formated_table)):
        if not formated_table[i][1] in chain(*deduplicated_table):
            deduplicated_table.append(formated_table[i])

    pdf.close_file()

    return deduplicated_table


def write_final_excel(file_path, save_path, broj_vagona=None, otpremna_zelj_uprava=None, sifra_otpremnog_kol=None,
                      uputna_zelj_uprava=None, sifra_uputnog_kol=None, okvirni_opis_tereta=None,
                      duzina_vagona=None, tara_vagona=None, neto_vagona=None, rucno_kocena_tezina=None,
                      zracno_kocena_tezina=None, slovna_serija=None, broj_osovina=None, otpremni_kolodvor=None,
                      uputni_kolodvor=None, isprava=None):
    """ Function for writing final XLSM file. Keep Excel open.

    :param isprava:
    :param file_path:
    :param save_path:
    :param broj_vagona:
    :param otpremna_zelj_uprava:
    :param sifra_otpremnog_kol:
    :param uputna_zelj_uprava:
    :param sifra_uputnog_kol:
    :param okvirni_opis_tereta:
    :param duzina_vagona:
    :param tara_vagona:
    :param neto_vagona:
    :param rucno_kocena_tezina:
    :param zracno_kocena_tezina:
    :param slovna_serija:
    :param broj_osovina:
    :param otpremni_kolodvor:
    :param uputni_kolodvor:
    :return:
    """
    workbook = xw.Book(file_path)
    worksheet = workbook.sheets["Sheet1"]
    worksheet.range('H10').options(transpose=True).value = broj_vagona
    worksheet.range('J10').options(transpose=True).value = otpremna_zelj_uprava
    worksheet.range('L10').options(transpose=True).value = sifra_otpremnog_kol
    worksheet.range('N10').options(transpose=True).value = uputna_zelj_uprava
    worksheet.range('P10').options(transpose=True).value = sifra_uputnog_kol
    worksheet.range('V10').options(transpose=True).value = okvirni_opis_tereta
    worksheet.range('X10').options(transpose=True).value = duzina_vagona
    worksheet.range('Z10').options(transpose=True).value = tara_vagona
    worksheet.range('AB10').options(transpose=True).value = neto_vagona
    worksheet.range('AD10').options(transpose=True).value = rucno_kocena_tezina
    worksheet.range('AH10').options(transpose=True).value = zracno_kocena_tezina
    worksheet.range('AL10').options(transpose=True).value = slovna_serija
    worksheet.range('AP10').options(transpose=True).value = broj_osovina
    worksheet.range('AR10').options(transpose=True).value = otpremni_kolodvor
    worksheet.range('AT10').options(transpose=True).value = uputni_kolodvor
    worksheet.range('AY10').options(transpose=True).value = isprava

    workbook.save(save_path)
