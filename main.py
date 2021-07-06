#!/usr/bin/env python

import os.path
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import file_operations
import tksheet


class GUI:
    table = [[""] * 11 for i in range(13)]
    head = ['No.', 'br. vagona', 'serija', 'tara', 'duljina', 'rKM', 'KM', 'br. osovina', 'KM prazan', 'neto', 'bruto']

    def __init__(self, master):
        self.master = master
        master.title("Obrada")

        self.first_file_path = tk.StringVar()
        self.second_file_path = tk.StringVar()
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
        self.kocna_masa_var = tk.IntVar()

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
                                   column_width=92, empty_vertical=0, empty_horizontal=0)
        self.sheet.enable_bindings()
        self.sheet.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self.sheet.headers(newheaders=self.head, index=None, reset_col_positions=False, show_headers_if_not_sheet=True)

        # File picker frame
        self.pick_first_file_btn = ttk.Button(self.pick_frame, text="Odaberi PDF s popisom", width=25,
                                              command=self.pick_first_file_btn_click)
        self.pick_first_file_btn.grid(row=0, column=0, padx=5, pady=5)

        self.first_filepath_label = ttk.Label(self.pick_frame, text="Putanja:")
        self.first_filepath_label.grid(row=0, column=1, padx=(15, 5), pady=5)
        self.first_filepath_entry = ttk.Entry(self.pick_frame, width=50, textvariable=self.first_file_path)
        self.first_filepath_entry.bind("<Return>", self.enter_path_first_file)
        self.first_filepath_entry.grid(row=0, column=2, padx=(0, 5), pady=5)

        # Calculation frame
        self.tare_label = ttk.Label(self.calc_frame, text="Tara težina (tone):")
        self.tare_label.grid(row=0, column=3, padx=(5, 0), pady=5, sticky="e")
        self.tare_calc = ttk.Entry(self.calc_frame, width=16, textvariable=self.tare_entry)
        self.tare_calc.grid(row=0, column=4, pady=5)

        self.bruto_label = ttk.Label(self.calc_frame, text="Bruto težina (tone):")
        self.bruto_label.grid(row=1, column=3, padx=(5, 0), pady=5, sticky="e")
        self.bruto_calc = ttk.Entry(self.calc_frame, width=16, textvariable=self.bruto_entry)
        self.bruto_calc.grid(row=1, column=4, pady=5)

        self.neto_label = ttk.Label(self.calc_frame, text="Neto težina (tone):")
        self.neto_label.grid(row=2, column=3, padx=(5, 0), pady=5, sticky="e")
        self.neto_calc = ttk.Entry(self.calc_frame, width=16, textvariable=self.neto_entry)
        self.neto_calc.grid(row=2, column=4, pady=5)

        self.calc_btn = ttk.Button(self.calc_frame, text="Izračunaj",  command=self.calculate_weights_btn)
        self.calc_btn.grid(row=3, column=3, columnspan=2, padx=(5, 0), pady=5, sticky="nsew")

        self.otpremna_zelj_uprava_label = ttk.Label(self.calc_frame, text="Otpremna želj. uprava:")
        self.otpremna_zelj_uprava_label.grid(row=0, column=5, padx=(5, 0), pady=5, sticky="e")
        self.otpremna_zelj_uprava_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.otpremna_zelj_uprava)
        self.otpremna_zelj_uprava_entry.grid(row=0, column=6, padx=5, pady=5)

        self.sifra_otpremnog_kol_label = ttk.Label(self.calc_frame, text="Šifra otpremnog kol.:")
        self.sifra_otpremnog_kol_label.grid(row=1, column=5, padx=(5, 0), pady=5, sticky="e")
        self.sifra_otpremnog_kol_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.sifra_otpremnog_kol)
        self.sifra_otpremnog_kol_entry.grid(row=1, column=6, padx=5, pady=5)

        self.uputna_zelj_uprava_label = ttk.Label(self.calc_frame, text="Uputna želj. uprava:")
        self.uputna_zelj_uprava_label.grid(row=2, column=5, padx=(5, 0), pady=5, sticky="e")
        self.uputna_zelj_uprava_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.uputna_zelj_uprava)
        self.uputna_zelj_uprava_entry.grid(row=2, column=6, padx=5, pady=5)

        self.sifra_uputnog_kol_label = ttk.Label(self.calc_frame, text="Šifra uputnog kol.:")
        self.sifra_uputnog_kol_label.grid(row=3, column=5, padx=(5, 0), pady=5, sticky="e")
        self.sifra_uputnog_kol_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.sifra_uputnog_kol)
        self.sifra_uputnog_kol_entry.grid(row=3, column=6, padx=5, pady=5)

        self.okvirni_opis_tereta_label = ttk.Label(self.calc_frame, text="Okvirni opis tereta:")
        self.okvirni_opis_tereta_label.grid(row=0, column=7, padx=(5, 0), pady=5, sticky="e")
        self.okvirni_opis_tereta_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.okvirni_opis_tereta)
        self.okvirni_opis_tereta_entry.grid(row=0, column=8, padx=5, pady=5)

        self.otpremni_kol_label = ttk.Label(self.calc_frame, text="Otpremni kolodvor:")
        self.otpremni_kol_label.grid(row=1, column=7, padx=(5, 0), pady=5, sticky="e")
        self.otpremni_kol_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.otpremni_kol)
        self.otpremni_kol_entry.grid(row=1, column=8, padx=5, pady=5)

        self.uputni_kol_label = ttk.Label(self.calc_frame, text="Uputni kolodvor:")
        self.uputni_kol_label.grid(row=2, column=7, padx=(5, 0), pady=5, sticky="e")
        self.uputni_kol_entry = ttk.Entry(self.calc_frame, width=16, textvariable=self.uputni_kol)
        self.uputni_kol_entry.grid(row=2, column=8, padx=5, pady=5)

        self.zracno_kocna_label = ttk.Label(self.calc_frame, text="Kočna masa:")
        self.zracno_kocna_label.grid(row=3, column=7, padx=(5, 0), pady=5, sticky="e")
        self.zracno_kocna_rbtn1 = ttk.Radiobutton(self.calc_frame, text="Pun", variable=self.kocna_masa_var, value=0)
        self.zracno_kocna_rbtn1.grid(row=3, column=8, padx=(2, 40), pady=5, sticky="w")
        self.zracno_kocna_rbtn2 = ttk.Radiobutton(self.calc_frame, text="Prazan", variable=self.kocna_masa_var, value=1)
        self.zracno_kocna_rbtn2.grid(row=3, column=8,  padx=(40, 0), pady=5, sticky="e")

        self.send_btn = ttk.Button(self.calc_frame, text="SEND", command=self.send_data)
        self.send_btn.grid(row=0, column=9, rowspan=4, sticky="nsew")
        self.calc_frame.grid_columnconfigure(9, weight=5)

    def pick_first_file_btn_click(self):
        path = filedialog.askopenfilename()
        if path != '':
            self.first_file_path.set(path)
            self.table = file_operations.open_first_pdf(self.first_file_path.get())
            self.sheet.set_sheet_data(self.table, reset_col_positions=True, reset_row_positions=True,
                                      redraw=True, verify=False, reset_highlights=False)
            self.sheet.align(align="center", redraw=True)

    def enter_path_first_file(self, event):
        if os.path.isfile(self.first_file_path.get()):
            self.table = file_operations.open_first_pdf(self.first_file_path.get())
            self.sheet.set_sheet_data(self.table, reset_col_positions=True, reset_row_positions=True,
                                      redraw=True, verify=False, reset_highlights=False)
            self.sheet.align(align="center", redraw=True)
        else:
            self.first_file_path.set("Molim unesite putanju do ispravne datoteke!")

    def calculate_weights_btn(self):
        result = self.calculate_weights(3)
        if result == "Greška!" or result == "":
            self.tare_entry.set(result)
        else:
            self.tare_entry.set(f'{result:.4f}')

        result = self.calculate_weights(9)
        if result == "Greška!":
            self.neto_entry.set(result)
        else:
            self.neto_entry.set(f'{result:.4f}')

        result = self.calculate_weights(10)
        if result == "Greška!":
            self.bruto_entry.set(result)
        else:
            self.bruto_entry.set(f'{result:.4f}')

    def calculate_weights(self, column):
        total = 0.0
        elements = self.sheet.get_column_data(column, return_copy=True)
        for i in range(len(elements)):
            elements[i].replace(',', '.')
        for el in elements:
            try:
                total = total + float(el)
            except ValueError:
                return "Greška!"
        return total

    def send_data(self):
        #test
        # file_operations.write_final_excel("Najava.xlsm", "Najava2.xlsm", self.sheet.get_column_data(1), self.sheet.get_column_data(2))
        pass

if __name__ == "__main__":
    root = tk.Tk()
    root.minsize(1110, 585)
    gui = GUI(root)
    root.mainloop()
