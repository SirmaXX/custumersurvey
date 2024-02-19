import tkinter as tk
from tkinter import messagebox 
from datetime import datetime
import sqlite3
import time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import logging
import os
import openpyxl

class RoomForm:
    def __init__(self, root):
        self.root = root
        self.root.title("Costumer Survey Example Project")
 
        
         # Create database and table
        self.connection = sqlite3.connect("crud.db")
        self.cursor = self.connection.cursor()
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS records (
                                id INTEGER PRIMARY KEY AUTOINCREMENT,
                                room INTEGER NOT NULL,
                                costumer TEXT NOT NULL,
                                giristime TEXT NOT NULL,
                                cikistime TEXT NOT NULL,
                                olumlu TEXT NOT NULL,
                                olumsuz TEXT NOT NULL,
                                comment TEXT NOT NULL)''')
        self.connection.commit()




        self.room_entry = None
        # Variables to store user input
        self.room_number_var = tk.IntVar()
        self.costumer_var = tk.StringVar()
        self.gtime_var = tk.StringVar()
        self.ctime_var = tk.StringVar()
        
        self.positive_var = tk.StringVar()
        self.negative_var = tk.StringVar()
        self.comment_var = tk.StringVar()
      
        # Create form elements
        self.create_widgets()



    def show_info_popup(self, message):
        messagebox.showinfo("İnfo", message) 
        
    def show_error_popup(self, message):
        messagebox.showerror("Error", message) 
    
    
    def open_popup(self):
     top= tk.Toplevel(self.root)
     top.geometry("550x150")
     top.title("Support Page")
     tk.Label(top, text= "Deniz Balcı +905536964026  sea-97@hotmail.com ").place(x=150,y=80)

    def select_item(self, event):
     try:
        global selected_item
        index = self.tree.focus()
        selected_item = self.tree.item(index, 'values')

        self.room_number_var.set(selected_item[1])
        self.gtime_var.set(selected_item[3])
        self.ctime_var.set(selected_item[4])
        self.costumer_var.set(selected_item[2])
        self.positive_var.set(selected_item[5])
        self.negative_var.set(selected_item[6])
        self.comment_var.set(selected_item[7])
        
        
        
        

     except IndexError:
        logging.error(IndexError)
        pass





    def create_widgets(self):
        # Room Number
        room_label = tk.Label(self.root, text="Oda Numarası:")
        room_entry = tk.Entry(self.root, textvariable=self.room_number_var)
        #Costumer
        costumer_label = tk.Label(self.root, text="Müşteri Adı Soyadı")
        costumer_entry = tk.Entry(self.root, textvariable=self.costumer_var)
        # Time
        time_label = tk.Label(self.root, text="Giriş Tarihi:")
        time_entry = tk.Entry(self.root, textvariable=self.gtime_var)
        self.set_default_time()
        
        
        
        time_label_1 = tk.Label(self.root, text="Çıkış Tarih:")
        time_entry_1= tk.Entry(self.root, textvariable=self.ctime_var)
        self.set_default_time_k()
        
        

      

        # Comment
        positive_comment_label = tk.Label(self.root, text="Olumlu")
        positive_comment_entry = tk.Entry(self.root, textvariable=self.positive_var)
        
        negative_comment_label = tk.Label(self.root, text="Olumsuz")
        negative_comment_entry = tk.Entry(self.root, textvariable=self.negative_var)
        
        improve_comment_label = tk.Label(self.root, text="Geliştirilecek yönler")
        improve_comment_entry = tk.Entry(self.root, textvariable=self.comment_var)
        
        

        # Buttons
        submit_button = tk.Button(self.root, text="Kayıt ekle", command=self.submit_form)
        view_button = tk.Button(self.root, text="Kayıtları gör", command=self.view_records)
        export_button = tk.Button(self.root, text="Excel çıktısı", command=self.excel_export)
        update_button = tk.Button(self.root, text="Kayıt güncelleme", command=self.update_record)
        delete_button = tk.Button(self.root, text="Kayıt sil", command=self.delete_record)
     
        support_button = tk.Button(self.root, text= "Destek sayfası", command= self.open_popup)
        pdf_button = tk.Button(self.root, text= "Pdf çıktısı", command= self.pdf_export)

        # Grid layout
        ## ilk satır
        room_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        room_entry.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        costumer_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        costumer_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        time_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        time_entry.grid(row=1, column=2, padx=10, pady=5, sticky="w")
        
        time_label_1.grid(row=0, column=3, padx=10, pady=5, sticky="w")
        time_entry_1.grid(row=1, column=3, padx=10, pady=5, sticky="w")

          ## ilk satır
        ## ikinci satır
        positive_comment_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        positive_comment_entry.grid(row=3, column=0, padx=10, pady=5, sticky="w")
      
      
        negative_comment_label.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        negative_comment_entry.grid(row=3, column=1, padx=10, pady=5, sticky="w")
      
      
        improve_comment_label.grid(row=2, column=2, padx=10, pady=5, sticky="w")
        improve_comment_entry.grid(row=3, column=2, padx=10, pady=5, sticky="w")
         ## ikinci satır    
        
      
          ## ikinci satır
        submit_button.grid(row=4, column=0, pady=10)
        view_button.grid(row=4, column=1, pady=10)
        export_button.grid(row=4, column=2, pady=10)
        
        update_button.grid(row=5, column=0, pady=10)
        delete_button.grid(row=5, column=1, pady=10)
        support_button.grid(row=5, column=2, pady=10)
        pdf_button.grid(row=4, column=3, pady=10)
        # Treeview
        columns = ("ID", "room", "costumer", "giristime","cikistime" ,"olumlu", "olumsuz","comment")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="browse")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")

        self.tree.grid(row=6, column=0, columnspan=3, pady=20, padx=20, sticky="nsew")

        # Create scrollbar
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=6, column=3, sticky="ns")

        # Set scroll to treeview
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Bind select
        self.tree.bind('<<TreeviewSelect>>', self.select_item)

    def set_default_time(self):
     # Set default time to current date and time
     current_datetime = datetime.now()
     current_time = current_datetime.strftime("%H:%M:%S")
     current_date = current_datetime.strftime("%d/%m/%Y")
     self.gtime_var.set(f"{current_date} - {current_time}")
     
     
    def set_default_time_k(self):
     # Set default time to current date and time
     current_datetime = datetime.now()
     current_time = current_datetime.strftime("%H:%M:%S")
     current_date = current_datetime.strftime("%d/%m/%Y")
     self.ctime_var.set(f"{current_date} - {current_time}")
      

    def submit_form(self):
     try:
        # Retrieve user inputs
        room_number = self.room_number_var.get()
        gtime_str = self.gtime_var.get()
        ctime_str = self.ctime_var.get()
        costumer = self.costumer_var.get()
        pcomment = self.positive_var.get()
        ncomment = self.negative_var.get()
        comment = self.comment_var.get()

        # Convert gtime and ctime strings to datetime objects
        gtime = datetime.strptime(gtime_str, "%d/%m/%Y - %H:%M:%S")
        ctime = datetime.strptime(ctime_str, "%d/%m/%Y - %H:%M:%S")

        if room_number and costumer and gtime and ctime and comment:
            # Check if gtime is greater than ctime
            if gtime > ctime:
                self.show_error_popup("Giriş Tarihi Çıkış Tarihinden sonra Olamaz")
                return

            self.cursor.execute("INSERT INTO records (room, costumer, giristime, cikistime, olumlu, olumsuz, comment) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                (room_number, costumer, gtime_str, ctime_str, pcomment, ncomment, comment))
            self.connection.commit()
            logging.info('Record added: Room Number: %s, Giriş Time: %s, Çıkış Time: %s, Costumer: %s, Olumlu: %s, Olumsuz: %s, Comment: %s',
                         room_number, gtime_str, ctime_str, costumer, pcomment, ncomment, comment)
            self.show_info_popup("Kayıt eklendi")
        else:
            self.show_error_popup("Lütfen gerekli bilgileri doldurunuz")
            logging.error('Failed to add record: Required information is missing')
     except ValueError:
        self.show_error_popup("Tarih formatı yanlış")
     except Exception as e:
        logging.exception('An error occurred while submitting the form: %s', e)
             
             
    def pdf_export(self):
     self.cursor.execute("SELECT * FROM records")
     records = self.cursor.fetchall()
     columns = ["ID", "Oda Numarası", "Giriş Tarihi", "Çıkış Tarihi", "Müşteri Adı Soyadı" , "Olumlu", "Olumsuz", "Geliştirilecek Yönler"]
    
     record_values = []
     for record in records:
        # Ensure the order of values matches the order of columns in the DataFrame
        record_values.append((record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7]))
    
     df = pd.DataFrame(record_values, columns=columns)
     fig, ax = plt.subplots(figsize=(12, 4))
     ax.axis('tight')
     ax.axis('off')
     the_table = ax.table(cellText=df.values, colLabels=df.columns, loc='center')
     pp = PdfPages("database.pdf")
     pp.savefig(fig, bbox_inches='tight')
     pp.close()
     self.show_info_popup("PDF dosyası oluşturuldu")
     logging.info('PDF file created')





    def excel_export(self):
     self.cursor.execute("SELECT * FROM records")
     records = self.cursor.fetchall()
     columns = ["ID", "Oda Numarası", "Giriş Tarihi", "Çıkış Tarihi", "Müşteri Adı Soyadı" , "Olumlu", "Olumsuz", "Geliştirilecek Yönler"]
    
     record_values = []
     for record in records:
        # Ensure the order of values matches the order of columns in the DataFrame
        record_values.append((record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7]))
    
     df = pd.DataFrame(record_values, columns=columns)
     df.to_excel('output.xlsx', index=False)
     self.show_info_popup("Excel dosyası oluşturuldu")
     logging.info('Excel file created')
        

    def view_records(self):
    # Clear existing items in the tree
     for record in self.tree.get_children():
        self.tree.delete(record)

    # Fetch records from the database
     self.cursor.execute("SELECT * FROM records")
     records = self.cursor.fetchall()
     print(records)
     for record in records:
        # Ensure the order of values matches the order of columns in the Treeview
        record_values = (record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7])
        self.tree.insert("", "end", values=record_values)
        
        

    def update_record(self):
        # Fetch the selected item's ID
        selected_id = selected_item[0] if 'selected_item' in globals() else None

        if selected_id:
            # Retrieve user inputs
            room_number = self.room_number_var.get()
            gtime = self.gtime_var.get()
            ctime = self.ctime_var.get()
            costumer = self.costumer_var.get()
            pcomment = self.positive_var.get()
            ncomment = self.negative_var.get()
            comment = self.comment_var.get()
            print(selected_id, room_number, costumer, gtime, ctime, costumer, pcomment, ncomment, comment)

            # Update the record in the database
            self.cursor.execute("UPDATE records SET room=?, giristime=?,costumer=?, cikistime=?, olumlu=?, olumsuz=?, comment=? WHERE id=?",
                                (room_number, gtime,costumer, ctime, pcomment, ncomment, comment, selected_id))
            self.connection.commit()
            self.show_info_popup("kayıt güncellendi")

            # Refresh the view after updating
            self.view_records()




    def delete_record(self):
     try:
        # Fetch the selected item's ID
        selected_id = selected_item[0] if 'selected_item' in globals() else None

        if selected_id:
            # Delete the record from the database
            self.cursor.execute("DELETE FROM records WHERE id=?", (selected_id,))
            self.connection.commit()
            self.show_info_popup("kayıt silindi")

            # Refresh the view after deleting
            self.view_records()
     except Exception as e:
        logging.exception('An error occurred while deleting the record: %s', e)





if __name__ == "__main__":
    logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    root = ttk.Window(themename="darkly")
    app = RoomForm(root)
    root.mainloop()