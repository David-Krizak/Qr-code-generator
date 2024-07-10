import tkinter as tk
from tkinter import messagebox
import qrcode
import win32com.client
import pythoncom
import io
import os
import tempfile

def generate_vcard(name, surname, company, email):
    vcard = f"""BEGIN:VCARD
VERSION:3.0
N:{surname};{name}
FN:{name} {surname}
ORG:{company}
EMAIL;WORK;INTERNET:{email}
END:VCARD"""
    return vcard

def generate_qr_code(vcard_data):
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(vcard_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    return img

def save_qr_code_image(qr_image):
    temp_dir = tempfile.gettempdir()
    temp_filename = os.path.join(temp_dir, "QRCode.png")
    qr_image.save(temp_filename)
    return temp_filename

def create_outlook_email(receiver_email, subject, body, qr_image_path):
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = receiver_email
        message.Subject = subject
        message.Body = body

        # Attach the file
        message.Attachments.Add(qr_image_path)

        message.Display()
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        pythoncom.CoUninitialize()
        return str(e)

class QRCodeGenerator:
    def __init__(self, master):
        self.master = master
        master.title("Generator QR koda za vCard")

        # Create input fields
        self.create_input_field("Ime:", 0)
        self.create_input_field("Prezime:", 1)
        self.create_input_field("Tvrtka:", 2)
        self.create_input_field("Email:", 3)

        # Create generate button
        self.generate_button = tk.Button(master, text="Generiraj QR kod i otvori u Outlooku", command=self.generate_and_open_outlook)
        self.generate_button.grid(row=4, column=0, columnspan=2, pady=10)

    def create_input_field(self, label, row):
        tk.Label(self.master, text=label).grid(row=row, column=0, sticky="e", padx=5, pady=5)
        entry = tk.Entry(self.master)
        entry.grid(row=row, column=1, padx=5, pady=5)
        setattr(self, label.lower().replace(":", "")+"_entry", entry)

    def generate_and_open_outlook(self):
        ime = self.ime_entry.get()
        prezime = self.prezime_entry.get()
        tvrtka = self.tvrtka_entry.get()
        email = self.email_entry.get()

        if not all([ime, prezime, tvrtka, email]):
            messagebox.showerror("Greška", "Sva polja su obavezna!")
            return

        vcard_data = generate_vcard(ime, prezime, tvrtka, email)
        qr_image = generate_qr_code(vcard_data)
        save_path = save_qr_code_image(qr_image)

        subject = "Vaš vCard QR kod"
        body = "U prilogu se nalazi vaš qr kod za pristup."

        result = create_outlook_email(email, subject, body, save_path)

        if result is True:
            messagebox.showinfo("Uspjeh", "QR kod je generiran i Outlook email je uspješno kreiran!")
        else:
            messagebox.showerror("Greška", f"Dogodila se greška:\n{result}")

def main():
    root = tk.Tk()
    QRCodeGenerator(root)
    root.geometry("300x200")
    root.mainloop()


if __name__ == "__main__":
    main()
