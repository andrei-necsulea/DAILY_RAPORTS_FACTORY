import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkcalendar as calendar
from tkinter import messagebox
import os
import xlsxwriter



window = tk.Tk()
window.title("RAPORT ZILNIC")
window.geometry("450x680")
window.resizable(0,0)
window.iconbitmap("iconita.ico")

font_var = ("Times New Roman" , 12 )
font_var_cmb_list = ("Times New Roman" , 10)
font_title_var = ("Times New Roman" , 20 , "bold")

title_var = tk.Label(text = " \n RAPORT  ZILNIC  MENTENANTA   \n")
title_var.place(relx = 0 , rely = 0)
title_var.config(font=font_title_var , bg="lightgray")

nume = tk.Label(text = "NUME : " )
nume.place(relx = 0.05 , rely = 0.2)
nume.config(font=font_var)

data = tk.Label(text = "DATA : ")
data.place(relx = 0.555 , rely = 0.2)
data.config(font=font_var)

tura  = tk.Label(text = "TURA : ")
tura.place(relx = 0.05 , rely = 0.32)
tura.config(font=font_var)

evob = tk.Label(text = "EVENIMENTE / OBIECTII : ")
evob.place(relx = 0.1 , rely = 0.45)
evob.config(font=font_var)

evob_textarea = tk.Text(height = 12 , font=("Times New Roman" , 12) , width = 50 , wrap=WORD)
evob_textarea.place(relx = 0.05 , rely=0.52)
evob.config(border=1)



nume_var = tk.StringVar()
nume_cmb = ttk.Combobox(font = font_var , width = 16 , textvariable=nume_var)
nume_cmb.place(rely = 0.2 , relx = 0.19)
nume_cmb['values'] = ("CIOCANETE MARIAN" , "MILITARU MARIAN" , "CALARASU CRISTI" , "TANTAREA IULIAN" , "CIUCEA STELICA" , "GHEORGHISOR CONSTANTIN")

def nume_func(event):
    global nume_aux
    nume_aux = nume_var.get()
    
nume_cmb.bind('<<ComboboxSelected>>', nume_func)


data_cmb = calendar.DateEntry(selectmode = "day" , width = 12 , date_pattern="dd.mm.yyyy" , locale = "ro_RO")
data_cmb.place(rely = 0.2 , relx = 0.69)
data_cmb.config(font=font_var)



def data_func(event):

    global data_aux
    data_aux = str(data_cmb.get_date())

    data_aux = data_aux.replace("-" , ".")

    year = data_aux[:4]
    day = data_aux[8:10]
    month = data_aux[5:7]

    data_aux = day + "." + month + "." + year                        
                       

data_cmb.bind("<<DateEntrySelected>>", data_func)



tura_var = tk.StringVar()
tura_cmb = ttk.Combobox(font = font_var , width = 4 , textvariable=tura_var)
tura_cmb.place(rely = 0.32 , relx = 0.19)
tura_cmb['values'] = ("1" , "2")

def tura_func(event):
    global tura_aux
    tura_aux = tura_var.get()
    
tura_cmb.bind('<<ComboboxSelected>>', tura_func)



window.option_add('*TCombobox*Listbox.font', font_var_cmb_list)
    

def creeaza_xlsx():

  global evenimente_edit
  k = 0
  try :

       current_user = os.getlogin()
       
       path = "C:\\Users\\" + current_user + "\\Desktop\\Rapoarte_Zilnice\\"
       
       excel_file_name =  path + data_aux + "_T" + tura_aux + ".xlsx"

       
       str_message = "CREATI FISIERUL EXCEL CU NUMELE {excel_filename_without_path} ?".format(excel_filename_without_path = data_aux + "_T" + tura_aux + ".xlsx")
       

       evenimente = evob_textarea.get("1.0",END)
       
       if messagebox.askokcancel("CREARE",str_message):
        workbook = xlsxwriter.Workbook(excel_file_name)

        worksheet1 = workbook.add_worksheet()

        worksheet1.write(2, 0, 'NUME : ')   
       
        worksheet1.write(5 , 4 , 'TURA : ')   
        worksheet1.write(2 , 6 , "DATA : ")

        merge_format = workbook.add_format({
           'bold': 1,
           'border': 1,
           'align': 'center',
           'valign': 'vcenter',
           'fg_color': '#D7E4BC',
           'text_wrap': True})

        worksheet1.merge_range('A9:I31', evenimente , merge_format)
       
        worksheet1.merge_range('B3:D3', nume_aux , merge_format)

        worksheet1.merge_range('H3:I3' , data_aux , merge_format)

        worksheet1.write("F6" , tura_aux , merge_format) 

        workbook.close()

        k = 1

       
       tura_aux_k = tura_aux
       data_aux_k = data_aux
       nume_aux_k = nume_aux
       
       
       evenimente_edit = evenimente
       
       
       
       
  except NameError :
       messagebox.showerror("EROARE","Una din casetele NUME , TURA , DATA este goala !\nSelectati ceva in caseta !")
  
  
  if k == 1 :

    create_btn.destroy()
    evob_textarea.delete("1.0", END)
    nume_cmb.config(state="disabled")
    tura_cmb.config(state= "disabled")
    data_cmb.config(state= "disabled")


    window.title("EDITARE RAPORT ZILNIC")
    title_var = tk.Label(text = "\nEDITARE RAPORT MENTENANTA\n")
    title_var.place(relx = 0 , rely = 0)
    title_var.config(font=font_title_var , bg="lightgray")

    def editare_btn():

      global evenimente_edit

      if messagebox.askokcancel("EDITARE","EDITATI FISIERUL ?"):
        evenimente_edit = evenimente_edit + evob_textarea.get("1.0",END) 
       

        workbook = xlsxwriter.Workbook(excel_file_name)

        worksheet1 = workbook.add_worksheet()

        worksheet1.write(2, 0, 'NUME : ')   

        worksheet1.write(5 , 4 , 'TURA : ')   
        worksheet1.write(2 , 6 , "DATA : ")

        merge_format = workbook.add_format({
           'bold': 1,
           'border': 1,
           'align': 'center',
           'valign': 'vcenter',
           'fg_color': '#D7E4BC',
           'text_wrap': True})

       

        worksheet1.merge_range('A9:I31', evenimente_edit , merge_format)
       
        worksheet1.merge_range('B3:D3', nume_aux_k , merge_format)

        worksheet1.merge_range('H3:I3' , data_aux_k , merge_format)

        worksheet1.write("F6" , tura_aux_k , merge_format) 


        while True:
          try:
            workbook.close()
            evob_textarea.delete("1.0", END)

          except xlsxwriter.exceptions.FileCreateError as e:
            decision =  messagebox.showwarning("EXCEL DESCHIS" , "INCHIDETI APLICATIA EXCEL !")

            if decision :
              continue

          break
       

      
    save_btn = tk.Button(text = "   ADAUGA DATELE  " , command=editare_btn)
    save_btn.place(relx = 0.285 , rely = 0.9)
    save_btn.config(font=font_var , bg="lightgray")


create_btn = tk.Button(text = "     CREEAZA EXCEL     " , command=creeaza_xlsx)
create_btn.place(relx = 0.285 , rely = 0.9)
create_btn.config(font=font_var , bg = "lightgray")


def on_closing():
    if messagebox.askokcancel("Iesiti", "IESITI DIN APLICATIE ?"):
        window.destroy()

window.protocol("WM_DELETE_WINDOW", on_closing)

window.mainloop()