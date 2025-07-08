import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import pywhatkit
import time
import threading
import os
import re

class WhatsAppBotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title('WhatsApp Excel Mesaj Botu')
        self.contacts_file = ''
        self.message_file = ''
        self.contacts_df = None
        self.message = ''
        self.setup_ui()

    def setup_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(padx=10, pady=10)

        tk.Button(frame, text='Excel Dosyası Seç', command=self.select_contacts).grid(row=0, column=0, padx=5, pady=5)
        self.contacts_label = tk.Label(frame, text='Seçilmedi')
        self.contacts_label.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(frame, text='Mesaj Dosyası Seç', command=self.select_message).grid(row=1, column=0, padx=5, pady=5)
        self.message_label = tk.Label(frame, text='Seçilmedi')
        self.message_label.grid(row=1, column=1, padx=5, pady=5)

        tk.Button(frame, text='Numaraları Yükle', command=self.load_contacts).grid(row=2, column=0, padx=5, pady=5)
        self.load_label = tk.Label(frame, text='')
        self.load_label.grid(row=2, column=1, padx=5, pady=5)

        self.contacts_text = scrolledtext.ScrolledText(self.root, width=50, height=10)
        self.contacts_text.pack(padx=10, pady=5)

        tk.Button(self.root, text='Mesajları Gönder', command=self.start_sending).pack(pady=10)
        # Selenium ile ilgili start_sending_selenium ve send_messages_selenium_gui fonksiyonları kaldırıldı

        self.log_text = scrolledtext.ScrolledText(self.root, width=50, height=10, state='disabled')
        self.log_text.pack(padx=10, pady=5)

    def select_contacts(self):
        file = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if file:
            self.contacts_file = file
            self.contacts_label.config(text=os.path.basename(file))

    def select_message(self):
        file = filedialog.askopenfilename(filetypes=[('Text Files', '*.txt')])
        if file:
            self.message_file = file
            self.message_label.config(text=os.path.basename(file))

    def load_contacts(self):
        if not self.contacts_file:
            messagebox.showerror('Hata', 'Lütfen bir Excel dosyası seçin.')
            return
        try:
            df = pd.read_excel(self.contacts_file)
            if 'Telefon' not in df.columns or 'Firma Adı' not in df.columns:
                messagebox.showerror('Hata', 'Excel dosyasında "Firma Adı" ve "Telefon" sütunları bulunmalı.')
                return
            # Sadece 05, +905 veya 5 ile başlayan numaraları al
            self.contacts_df = df[df['Telefon'].notna() & (
                df['Telefon'].astype(str).str.strip().str.startswith('05') |
                df['Telefon'].astype(str).str.strip().str.startswith('+905') |
                df['Telefon'].astype(str).str.strip().str.startswith('5')
            )].copy()
            def format_number(x):
                x = str(x)
                x = re.sub(r'\D', '', x)  # Sadece rakamlar kalsın
                if x.startswith('0'):
                    x = x[1:]
                if not x.startswith('90'):
                    x = '90' + x
                return x
            self.contacts_df['Telefon'] = self.contacts_df['Telefon'].astype(str).apply(format_number)
            self.contacts_text.delete('1.0', tk.END)
            for idx, row in self.contacts_df.iterrows():
                self.contacts_text.insert(tk.END, f"{row['Firma Adı']} - {row['Telefon']}\n")
            self.load_label.config(text=f"{len(self.contacts_df)} numara yüklendi.")
        except Exception as e:
            messagebox.showerror('Hata', f'Excel dosyası okunamadı: {e}')

    def load_message(self):
        if not self.message_file:
            messagebox.showerror('Hata', 'Lütfen bir mesaj dosyası seçin.')
            return False
        try:
            with open(self.message_file, 'r', encoding='utf-8') as f:
                self.message = f.read().strip()
            return True
        except Exception as e:
            messagebox.showerror('Hata', f'Mesaj dosyası okunamadı: {e}')
            return False

    def start_sending(self):
        if self.contacts_df is None or self.contacts_df.empty:
            messagebox.showerror('Hata', 'Önce numaraları yükleyin.')
            return
        if not self.load_message():
            return
        threading.Thread(target=self.send_messages, daemon=True).start()

    def send_messages(self):
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log('Lütfen WhatsApp Web QR kodunu telefonunuzla tarayın...')
        time.sleep(10)
        for idx, row in self.contacts_df.iterrows():
            phone = str(row['Telefon'])
            name = str(row['Firma Adı'])
            # Mesaj şablonundaki {firma_adi} placeholder'ını işletme adıyla değiştir
            personalized_message = self.message.replace('{firma_adi}', name)
            self.log(f'Gönderilecek mesaj: {personalized_message}')
            try:
                pywhatkit.sendwhatmsg_instantly(f'+{phone}', personalized_message, wait_time=8, tab_close=True, close_time=5)
                self.log(f'{name} ({phone}) numarasına mesaj gönderildi.')
            except Exception as e:
                self.log(f'Hata: {name} ({phone}) numarasına mesaj gönderilemedi. {e}')
            self.root.update()
            time.sleep(3)  # 3 saniye bekle
        self.log('Mesajlar gönderildi.')
        self.log_text.config(state='disabled')

    def log(self, msg):
        self.log_text.insert(tk.END, msg + '\n')
        self.log_text.see(tk.END)

# Selenium ile ilgili start_sending_selenium ve send_messages_selenium_gui fonksiyonları kaldırıldı
# send_messages_selenium fonksiyonu kaldırıldı 

if __name__ == '__main__':
    root = tk.Tk()
    app = WhatsAppBotGUI(root)
    root.mainloop() 