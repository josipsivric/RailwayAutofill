#!/usr/bin/env python

import os.path
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import file_operations
import tksheet
import re


class GUI:
    table = [[""] * 18 for i in range(13)]
    widths = [30, 100, 60, 60, 60, 60, 70, 70, 70, 70, 60, 60, 60, 70, 50, 80, 80, 60]
    head = ['No.', 'br. vagona', 'otpremna\nželj. upr.', 'šifra\notp. kol.', 'uputna\nželj. upr.', 'šifra\nuput. kol.',
            'okvirni\nopis tereta', 'dužina\nvagona (m)', 'tara\nvagona (t)', 'neto\nvagona (t)', 'ručno\nKM',
            'pun\nKM', 'prazan\nKM', 'serija', 'broj\nosovina', 'otpremni\nkolodvor', 'uputni\nkolodvor', 'isprava']

    def __init__(self, master):
        self.master = master
        master.title("Obrada")
        self.style = ttk.Style()
        self.style.configure('big.TButton', font=(None, 32, 'bold'), foreground="red")
        self.style.configure('bold.TLabel', font=(None, 12, 'bold'))

        self.first_file_path = tk.StringVar()
        self.smjer = tk.IntVar()
        self.org_excel_file_path = tk.StringVar()
        self.new_excel_file_path = tk.StringVar()

        self.tare_entry = tk.StringVar()
        self.bruto_entry = tk.StringVar()
        self.neto_entry = tk.StringVar()

        self.otpremna_zelj_uprava = tk.StringVar()
        self.sifra_otpremnog_kol = tk.StringVar()
        self.uputna_zelj_uprava = tk.StringVar()
        self.sifra_uputnog_kol = tk.StringVar()
        self.okvirni_opis_tereta = tk.StringVar()
        self.otpremni_kol = tk.StringVar()
        self.uputni_kol = tk.StringVar()
        self.isprava = tk.StringVar()
        self.kocna_masa_var = tk.IntVar()

        self.kol_usputne_manip = tk.StringVar()
        self.kol_usputne_manip.set("0")
        self.sif_usputne_manip = tk.StringVar()
        self.sif_usputne_manip.set("0")
        self.vrsta_zracne_kocnice = tk.StringVar()
        self.vrsta_zracne_kocnice.set("P")

        # Root window geometry
        self.data_frame = ttk.LabelFrame(master, text="Podatci")
        self.data_frame.grid(row=0, column=0, padx=10, pady=5, sticky="nsew")
        master.grid_rowconfigure(0, weight=1)
        master.grid_columnconfigure(0, weight=1)

        self.pick_frame = ttk.LabelFrame(master, text="Datoteke")
        self.pick_frame.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        self.calc_frame = ttk.LabelFrame(master, text="Ostali podatci")
        self.calc_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")

        # Sheet frame
        self.sheet = tksheet.Sheet(self.data_frame, data=self.table, show_row_index=True, show_top_left=False,
                                   column_width=92, empty_vertical=0, empty_horizontal=0, header_height="2")
        self.sheet.enable_bindings()
        self.set_widths()
        self.sheet.headers(newheaders=self.head, index=None, reset_col_positions=False, show_headers_if_not_sheet=True)
        self.sheet.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        # File picker frame
        self.pick_first_file_btn = ttk.Button(self.pick_frame, text="Odaberi PDF s popisom", width=25,
                                              command=self.pick_first_file_btn_click)
        self.pick_first_file_btn.grid(row=0, column=0, padx=5, pady=5)

        self.first_filepath_label = ttk.Label(self.pick_frame, text="Putanja:")
        self.first_filepath_label.grid(row=0, column=1, padx=5, pady=5)
        self.first_filepath_entry = ttk.Entry(self.pick_frame, width=61, textvariable=self.first_file_path)
        self.first_filepath_entry.bind("<Return>", self.enter_path_first_file)
        self.first_filepath_entry.grid(row=0, column=2, padx=(0, 5), pady=5)

        self.smjer_datoteke_rbtn1 = ttk.Radiobutton(self.pick_frame, text="Normalno", variable=self.smjer, value=0)
        self.smjer_datoteke_rbtn1.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.smjer_datoteke_rbtn2 = ttk.Radiobutton(self.pick_frame, text="Naopako", variable=self.smjer, value=1)
        self.smjer_datoteke_rbtn2.grid(row=1, column=0, padx=5, pady=5, sticky="e")

        self.pick_excel_org_btn = ttk.Button(self.pick_frame, text="Odaberi EXCEL za najavu", width=25,
                                             command=self.pick_excel_org_btn_click)
        self.pick_excel_org_btn.grid(row=0, column=3, padx=(42, 5), pady=5)

        self.org_excel_file_path_label = ttk.Label(self.pick_frame, text="Putanja:")
        self.org_excel_file_path_label.grid(row=0, column=4, padx=5, pady=5)
        self.org_excel_file_path_entry = ttk.Entry(self.pick_frame, width=61, textvariable=self.org_excel_file_path)
        self.org_excel_file_path_entry.bind("<Return>", self.enter_path_org_excel_file)
        self.org_excel_file_path_entry.grid(row=0, column=5, padx=(0, 5), pady=5)

        self.new_excel_file_path_label = ttk.Label(self.pick_frame, text="Putanja do nove generirane datoteke:")
        self.new_excel_file_path_label.grid(row=1, column=3, columnspan=2, padx=(15, 5), pady=5, sticky="e")
        self.new_excel_file_path_entry = ttk.Entry(self.pick_frame, width=61, textvariable=self.new_excel_file_path)
        self.new_excel_file_path_entry.grid(row=1, column=5, padx=(0, 5), pady=5)

        # Calculation frame
        self.podloga1 = tk.Frame(self.calc_frame, background="lightgreen")
        self.podloga1.grid(row=0, column=0, rowspan=4, columnspan=2, sticky="nsew")

        self.tare_label = ttk.Label(self.calc_frame, text="Tara težina (tone):", background="lightgreen")
        self.tare_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.tare_calc = ttk.Entry(self.calc_frame, width=15, textvariable=self.tare_entry)
        self.tare_calc.grid(row=0, column=1, pady=5, padx=5)

        self.neto_label = ttk.Label(self.calc_frame, text="Neto težina (tone):", background="lightgreen")
        self.neto_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.neto_calc = ttk.Entry(self.calc_frame, width=15, textvariable=self.neto_entry)
        self.neto_calc.grid(row=1, column=1, pady=5, padx=5)

        self.bruto_label = ttk.Label(self.calc_frame, text="Bruto težina (tone):", background="lightgreen")
        self.bruto_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.bruto_calc = ttk.Entry(self.calc_frame, width=15, textvariable=self.bruto_entry)
        self.bruto_calc.grid(row=2, column=1, pady=5, padx=5)

        self.calc_btn = ttk.Button(self.calc_frame, text="Izračunaj", command=self.calculate_weights_btn)
        self.calc_btn.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        self.otpremna_zelj_uprava_label = ttk.Label(self.calc_frame, text="Otpremna želj. uprava:")
        self.otpremna_zelj_uprava_label.grid(row=0, column=2, padx=(5, 0), pady=5, sticky="e")
        self.otpremna_zelj_uprava_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.otpremna_zelj_uprava)
        self.otpremna_zelj_uprava_entry.bind('<FocusOut>', lambda x: self.evaluate(2, self.otpremna_zelj_uprava.get()))
        self.otpremna_zelj_uprava_entry.bind('<Return>', lambda x: self.evaluate(2, self.otpremna_zelj_uprava.get()))
        self.otpremna_zelj_uprava_entry.grid(row=0, column=3, padx=5, pady=5)

        self.sifra_otpremnog_kol_label = ttk.Label(self.calc_frame, text="Šifra otpremnog kol.:")
        self.sifra_otpremnog_kol_label.grid(row=1, column=2, padx=(5, 0), pady=5, sticky="e")
        self.sifra_otpremnog_kol_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.sifra_otpremnog_kol)
        self.sifra_otpremnog_kol_entry.bind('<FocusOut>', lambda x: self.evaluate(3, self.sifra_otpremnog_kol.get()))
        self.sifra_otpremnog_kol_entry.bind('<Return>', lambda x: self.evaluate(3, self.sifra_otpremnog_kol.get()))
        self.sifra_otpremnog_kol_entry.grid(row=1, column=3, padx=5, pady=5)

        self.uputna_zelj_uprava_label = ttk.Label(self.calc_frame, text="Uputna želj. uprava:")
        self.uputna_zelj_uprava_label.grid(row=2, column=2, padx=(5, 0), pady=5, sticky="e")
        self.uputna_zelj_uprava_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.uputna_zelj_uprava)
        self.uputna_zelj_uprava_entry.bind('<FocusOut>', lambda x: self.evaluate(4, self.uputna_zelj_uprava.get()))
        self.uputna_zelj_uprava_entry.bind('<Return>', lambda x: self.evaluate(4, self.uputna_zelj_uprava.get()))
        self.uputna_zelj_uprava_entry.grid(row=2, column=3, padx=5, pady=5)

        self.sifra_uputnog_kol_label = ttk.Label(self.calc_frame, text="Šifra uputnog kol.:")
        self.sifra_uputnog_kol_label.grid(row=3, column=2, padx=(5, 0), pady=5, sticky="e")
        self.sifra_uputnog_kol_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.sifra_uputnog_kol)
        self.sifra_uputnog_kol_entry.bind('<FocusOut>', lambda x: self.evaluate(5, self.sifra_uputnog_kol.get()))
        self.sifra_uputnog_kol_entry.bind('<Return>', lambda x: self.evaluate(5, self.sifra_uputnog_kol.get()))
        self.sifra_uputnog_kol_entry.grid(row=3, column=3, padx=5, pady=5)

        self.okvirni_opis_tereta_label = ttk.Label(self.calc_frame, text="Okvirni opis tereta:")
        self.okvirni_opis_tereta_label.grid(row=0, column=4, padx=(5, 0), pady=5, sticky="e")
        self.okvirni_opis_tereta_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.okvirni_opis_tereta)
        self.okvirni_opis_tereta_entry.bind('<FocusOut>', lambda x: self.evaluate(6, self.okvirni_opis_tereta.get()))
        self.okvirni_opis_tereta_entry.bind('<Return>', lambda x: self.evaluate(6, self.okvirni_opis_tereta.get()))
        self.okvirni_opis_tereta_entry.grid(row=0, column=5, padx=5, pady=5)

        self.otpremni_kol_label = ttk.Label(self.calc_frame, text="Otpremni kolodvor:")
        self.otpremni_kol_label.grid(row=1, column=4, padx=(5, 0), pady=5, sticky="e")
        self.otpremni_kol_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.otpremni_kol)
        self.otpremni_kol_entry.bind('<FocusOut>', lambda x: self.evaluate(15, self.otpremni_kol.get()))
        self.otpremni_kol_entry.bind('<Return>', lambda x: self.evaluate(15, self.otpremni_kol.get()))
        self.otpremni_kol_entry.grid(row=1, column=5, padx=5, pady=5)

        self.uputni_kol_label = ttk.Label(self.calc_frame, text="Uputni kolodvor:")
        self.uputni_kol_label.grid(row=2, column=4, padx=(5, 0), pady=5, sticky="e")
        self.uputni_kol_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.uputni_kol)
        self.uputni_kol_entry.bind('<FocusOut>', lambda x: self.evaluate(16, self.uputni_kol.get()))
        self.uputni_kol_entry.bind('<Return>', lambda x: self.evaluate(16, self.uputni_kol.get()))
        self.uputni_kol_entry.grid(row=2, column=5, padx=5, pady=5)

        self.isprava_label = ttk.Label(self.calc_frame, text="Isprava:")
        self.isprava_label.grid(row=3, column=4, padx=(5, 0), pady=5, sticky="e")
        self.isprava_entry = ttk.Entry(self.calc_frame, width=20, textvariable=self.isprava)
        self.isprava_entry.bind('<FocusOut>', lambda x: self.evaluate(17, self.isprava.get()))
        self.isprava_entry.bind('<Return>', lambda x: self.evaluate(17, self.isprava.get()))
        self.isprava_entry.grid(row=3, column=5, padx=5, pady=5)

        self.zracno_kocna_label = ttk.Label(self.calc_frame, text="Kočna masa:")
        self.zracno_kocna_label.grid(row=0, column=6, padx=(5, 60), pady=5, sticky="w")
        self.zracno_kocna_rbtn1 = ttk.Radiobutton(self.calc_frame, text="Pun", variable=self.kocna_masa_var, value=0)
        self.zracno_kocna_rbtn1.grid(row=0, column=6, padx=5, pady=5, sticky="e")
        self.zracno_kocna_rbtn2 = ttk.Radiobutton(self.calc_frame, text="Prazan", variable=self.kocna_masa_var, value=1)
        self.zracno_kocna_rbtn2.grid(row=0, column=7, padx=5, pady=5, sticky="w")

        self.podloga = tk.Frame(self.calc_frame, background="lightgreen")
        self.podloga.grid(row=1, column=6, rowspan=4, columnspan=2, pady=(0, 5), sticky="nsew")

        # self.auto_label = ttk.Label(self.calc_frame, text="AUTOMATSKI", background="lightgreen", style='bold.TLabel')
        # self.auto_label.grid(row=0, column=6, columnspan=2, padx=5, pady=5)

        self.kol_usputne_manip_label = ttk.Label(self.calc_frame, text="Kolodvor usputne manipulacije: 0", background="lightgreen")
        self.kol_usputne_manip_label.grid(row=1, column=6, columnspan=2, padx=5, pady=5)

        self.sif_usputne_manip_label = ttk.Label(self.calc_frame, text="Šifra usputne manipulacije: 0", background="lightgreen")
        self.sif_usputne_manip_label.grid(row=2, column=6, columnspan=2, padx=5, pady=5)

        self.vrsta_zracne_koc_label = ttk.Label(self.calc_frame, text="Vrsta zračne kočnice: P", background="lightgreen")
        self.vrsta_zracne_koc_label.grid(row=3, column=6, columnspan=2, padx=5, pady=5)

        self.send_btn = ttk.Button(self.calc_frame, text="POŠALJI", command=self.send_data, style="big.TButton")
        self.send_btn.grid(row=0, column=8, rowspan=4, columnspan=4, padx=5, pady=(0, 5), sticky="nsew")
        self.calc_frame.grid_columnconfigure(8, weight=1)

    def pick_first_file_btn_click(self):
        """ Parsing file found via button.

        :return: None
        """
        path = filedialog.askopenfilename()
        if path != '':
            self.first_file_path.set(path)
            self.table = file_operations.open_first_pdf(self.first_file_path.get())
        if self.smjer.get() == 0:
            self.full_redraw_sheet(self.table)
        else:
            self.table.reverse()
            self.full_redraw_sheet(self.table)

    def enter_path_first_file(self, event):
        """ Parsing file via provided path.

        :param event:
        :return: None
        """
        if os.path.isfile(self.first_file_path.get()):
            self.table = file_operations.open_first_pdf(self.first_file_path.get())
            if self.smjer.get() == 0:
                self.full_redraw_sheet(self.table)
            else:
                self.table.reverse()
                self.full_redraw_sheet(self.table)
        else:
            self.first_file_path.set("Molim unesite putanju do ispravne datoteke!")

    def pick_excel_org_btn_click(self):
        """ Parsing file found via button.

        :return: None
        """
        path = filedialog.askopenfilename()
        if path != '':
            self.org_excel_file_path.set(path)
            self.enter_path_new_excel_file()

    def enter_path_org_excel_file(self, event):
        """ Parsing file via provided path.

        :param event:
        :return: None
        """
        if not os.path.isfile(self.org_excel_file_path.get()):
            self.first_file_path.set("Molim unesite putanju do ispravne datoteke!")

    def enter_path_new_excel_file(self):
        """ Parsing file via provided path.

        :return: None
        """
        directory = os.path.dirname(self.org_excel_file_path.get())
        filename = os.path.basename(self.org_excel_file_path.get())
        new = directory + "/new_" + filename
        self.new_excel_file_path.set(new)

    def evaluate(self, column, data):
        """ Evaluate data in entry boxes and fill sheet accordingly.

        :param column: Column which will be modified
        :param data: Data which will be evaluated. It should be in format VALUEdelimiterBEGINNINGdelimiterENDING.
                        delimiter is at least one empty space or at least one comma. All values are optional.
        :return: None
        """
        data = data.replace(' ', ',')
        new = re.sub("(?P<char>[" + re.escape(",") + "])(?P=char)+", r"\1", data)
        elem = new.split(',')
        if len(elem) == 1:
            for i in range(self.sheet.get_total_rows()):
                self.sheet.set_cell_data(i, column, value=elem[0], set_copy=True, redraw=True)
        elif len(elem) == 2:
            if int(elem[1]) >= self.sheet.get_total_rows():
                for i in range(0, self.sheet.get_total_rows()):
                    self.sheet.set_cell_data(i, column, value=elem[0], set_copy=True, redraw=True)
            else:
                for i in range(0, int(elem[1])):
                    self.sheet.set_cell_data(i, column, value=elem[0], set_copy=True, redraw=True)
        elif len(elem) >= 3:
            if int(elem[1]) > int(elem[2]):
                elem[1], elem[2] = elem[2], elem[1]
            if int(elem[2]) >= self.sheet.get_total_rows():
                for i in range(int(elem[1]), self.sheet.get_total_rows()):
                    self.sheet.set_cell_data(i, column, value=elem[0], set_copy=True, redraw=True)
            else:
                for i in range(int(elem[1]) - 1, int(elem[2])):
                    self.sheet.set_cell_data(i, column, value=elem[0], set_copy=True, redraw=True)

    def full_redraw_sheet(self, table):
        """ Redraw entire sheet and replace all values with new ones.

        :param table: New values to be inserted
        :return: None
        """
        self.sheet.set_sheet_data(table, reset_col_positions=True, reset_row_positions=True,
                                  redraw=True, verify=False, reset_highlights=False)
        self.sheet.align(align="center", redraw=True)
        self.set_widths()

    def set_widths(self):
        """ Format table to inimize whitespace and keep readability.

        :return: None
        """
        for col, width in enumerate(self.widths):
            self.sheet.column_width(column=col, width=width, only_set_if_too_small=False, redraw=True)

    def calculate_weights_btn(self):
        """ Call function to calculate total Tara, Netto, and Brutto weights on button click.

        :return: None
        """
        tara = self.calculate_weights(8)
        if tara == "Greška!" or tara == "":
            self.tare_entry.set(tara)
        else:
            self.tare_entry.set(f'{tara:.4f}')

        neto = self.calculate_weights(9)
        if neto == "Greška!" or neto == "":
            self.neto_entry.set(neto)
        else:
            self.neto_entry.set(f'{neto:.4f}')

        if tara == "Greška!" or neto == "Greška!" or tara == "" or neto == "":
            self.bruto_entry.set("Greška!")
        else:
            self.bruto_entry.set(f'{neto+tara:.4f}')

    def calculate_weights(self, column):
        """ Function for calculating totals in specific columns.

        :param column: Column to be calculated.
        :return: Total weight.
        """
        total = 0.0
        elements = self.sheet.get_column_data(column, return_copy=True)
        for i in range(len(elements)):
            elements[i].replace(',', '.')
        for el in elements:
            if el == '' or el == 'praznina':
                continue
            else:
                try:
                    total = total + float(el)
                except ValueError:
                    return "Greška!"
        return total

    def send_data(self):

        if self.kocna_masa_var.get() == 0:
            pun_prazan = self.sheet.get_column_data(11)
        else:
            pun_prazan = self.sheet.get_column_data(12)

        stupac = self.sheet.get_column_data(13)

        slovna_serija = []
        for i in range(len(stupac)):
            if stupac[i] != '':
                slovna_serija.append(stupac[i][0])
            else:
                slovna_serija.append('')

        stupac_isprava = self.sheet.get_column_data(17)
        isprava = []
        for i in range(len(stupac_isprava)):
            if stupac_isprava[i] != '':
                isprava.append('\'' + stupac_isprava[i])
            else:
                isprava.append('praznina')

        stupac_neto = self.sheet.get_column_data(9)
        neto = []
        for i in range(len(stupac_neto)):
            if stupac_neto[i] == '':
                neto.append('praznina')
            else:
                neto.append(stupac_neto[i])

        stupac_okvirni = self.sheet.get_column_data(6)
        okvirni_opis = []
        for i in range(len(stupac_okvirni)):
            if stupac_okvirni[i] == '':
                okvirni_opis.append('praznina')
            else:
                okvirni_opis.append(stupac_okvirni[i])

        file_operations.write_final_excel(self.org_excel_file_path.get(), self.new_excel_file_path.get(),
                                          self.sheet.get_column_data(1), self.sheet.get_column_data(2),
                                          self.sheet.get_column_data(3), self.sheet.get_column_data(4),
                                          self.sheet.get_column_data(5), okvirni_opis, self.sheet.get_column_data(7),
                                          self.sheet.get_column_data(8), neto, self.sheet.get_column_data(10),
                                          pun_prazan, slovna_serija, self.sheet.get_column_data(14),
                                          self.sheet.get_column_data(15), self.sheet.get_column_data(16), isprava)


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1270x635")
    root.minsize(1270, 635)
    gui = GUI(root)
    root.mainloop()
