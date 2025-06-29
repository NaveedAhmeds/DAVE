import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import os
from PIL import Image, ImageTk
from pdf2docx import Converter
from docx2pdf import convert

class DaveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("D.A.V.E. - Document Adapter & Versatile Encoder")
        self.root.geometry("800x600")
        self.root.resizable(False, False)

        # Set window icon
        try:
            self.logo = tk.PhotoImage(file="logo.png")
            self.root.iconphoto(False, self.logo)
        except Exception as e:
            print("Could not load app icon:", e)

        # Load and set background image
        self.bg_image = Image.open("bg.jpeg")
        self.bg_image = self.bg_image.resize((800, 600), Image.Resampling.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)

        self.background_label = tk.Label(root, image=self.bg_photo)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

        self.title_label = tk.Label(
            root,
            text="D.A.V.E.",
            font=("Georgia", 48, "bold italic"),
            bg="#f4eee8",
            fg="#5b4636"
        )
        self.title_label.place(relx=0.5, rely=0.12, anchor="center")

        # File display label
        self.file_label = tk.Label(
            root,
            text="",
            font=("Helvetica", 15, "bold"),
            bg="#f4eee8",
            fg="#5b4636"
        )
        self.file_label.place(relx=0.5, rely=0.4, anchor="center")

        # Buttons frame
        self.buttons_frame = tk.Frame(root, bg="#f4eee8")
        self.buttons_frame.place(relx=0.5, rely=0.3, anchor="center")

        self.btn_pdf_to_docx = ttk.Button(
            self.buttons_frame,
            text="Convert PDF to DOCX",
            command=self.select_pdf_file,
            width=25
        )
        self.btn_pdf_to_docx.pack(side=tk.LEFT, padx=20)

        self.btn_docx_to_pdf = ttk.Button(
            self.buttons_frame,
            text="Convert DOCX to PDF",
            command=self.select_docx_file,
            width=25
        )
        self.btn_docx_to_pdf.pack(side=tk.LEFT, padx=20)

        # Progress bar style and bar
        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure(
            "green.Horizontal.TProgressbar",
            thickness=25,
            troughcolor='#d3d3d3',
            background='green'
        )
        self.progress = ttk.Progressbar(
            root,
            orient=tk.HORIZONTAL,
            length=400,
            mode='determinate',
            style="green.Horizontal.TProgressbar",
            maximum=100
        )
        self.progress.place(relx=0.5, rely=0.45, anchor="center")
        self.progress.place_forget()

        self.progress_label = tk.Label(
            root,
            text="Progress: 0%",
            font=("Helvetica", 20, "bold"),
            bg="#f4eee8",
            fg="#5b4636"
        )
        self.progress_label.place(relx=0.5, rely=0.53, anchor="center")
        self.progress_label.place_forget()

        # Quote label
        self.quote_label = tk.Label(
            root,
            text='“Change in all things is sweet.” — Aristotle\nThe issue is not the change itself, but how we embrace and direct it.',
            font=("Georgia", 18, "bold italic"),
            bg="#f4eee8",
            fg="#5b4636",
            justify="center",
            wraplength=700,
            pady=10
        )
        self.quote_label.place(relx=0.5, rely=0.87, anchor="center")

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Welcome to D.A.V.E.! Select a conversion to start.")
        self.status_bar = tk.Label(
            root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor="w",
            bg="#dcd0c0",
            fg="#3e2e1c",
            font=("Arial", 15)
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf")],
        )
        if file_path:
            self.status_var.set(f"Selected PDF: {file_path}")
            self.file_label.config(text=f"Converting: {os.path.basename(file_path)}")
            self.start_conversion(file_path, "pdf_to_docx")

    def select_docx_file(self):
        file_path = filedialog.askopenfilename(
            title="Select DOCX File",
            filetypes=[("Word Documents", "*.docx")],
        )
        if file_path:
            self.status_var.set(f"Selected DOCX: {file_path}")
            self.file_label.config(text=f"Converting: {os.path.basename(file_path)}")
            self.start_conversion(file_path, "docx_to_pdf")

    def start_conversion(self, file_path, conversion_type):
        self.btn_pdf_to_docx.config(state=tk.DISABLED)
        self.btn_docx_to_pdf.config(state=tk.DISABLED)

        self.progress.place(relx=0.5, rely=0.45, anchor="center")
        self.progress_label.place(relx=0.5, rely=0.53, anchor="center")
        self.progress['value'] = 0
        self.progress_label.config(text="Progress: 0%")
        self.status_var.set("Conversion started...")

        threading.Thread(
            target=self.run_conversion,
            args=(file_path, conversion_type),
            daemon=True
        ).start()

    def run_conversion(self, file_path, conversion_type):
        try:
            base, ext = os.path.splitext(file_path)
            output_file = base + ("_daved.docx" if conversion_type == "pdf_to_docx" else "_daved.pdf")

            for i in range(101):
                time.sleep(0.015)
                self.root.after(0, self.update_progress, i)

            if conversion_type == "pdf_to_docx":
                self.convert_pdf_to_docx(file_path, output_file)
            else:
                self.convert_docx_to_pdf(file_path, output_file)
        except Exception as e:
            self.root.after(0, self.handle_error, str(e))
            return

        self.root.after(0, self.finish_conversion, output_file)

    def update_progress(self, value):
        self.progress['value'] = value
        self.progress_label.config(text=f"Progress: {value}%")

    def finish_conversion(self, output_file):
        self.progress.stop()
        self.progress.place_forget()
        self.progress_label.place_forget()
        self.progress['value'] = 0
        self.file_label.config(text="")

        self.btn_pdf_to_docx.config(state=tk.NORMAL)
        self.btn_docx_to_pdf.config(state=tk.NORMAL)

        self.status_var.set(f"Conversion completed: {output_file}")

        save_path = filedialog.asksaveasfilename(
            initialfile=os.path.basename(output_file),
            defaultextension=os.path.splitext(output_file)[1],
            filetypes=[("All files", "*.*")]
        )

        if save_path:
            try:
                os.replace(output_file, save_path)
                self.status_var.set(f"File saved to: {save_path}")
                messagebox.showinfo("Success", f"File saved to:\n{save_path}")
            except Exception as e:
                self.status_var.set("Failed to save file.")
                messagebox.showerror("Error", f"Failed to save file:\n{e}")
        else:
            self.status_var.set("Save canceled. Conversion file not saved.")

    def handle_error(self, error_msg):
        self.progress.stop()
        self.progress.place_forget()
        self.progress_label.place_forget()
        self.progress['value'] = 0
        self.file_label.config(text="")
        self.btn_pdf_to_docx.config(state=tk.NORMAL)
        self.btn_docx_to_pdf.config(state=tk.NORMAL)
        self.status_var.set("Conversion failed.")
        messagebox.showerror("Conversion Error", error_msg)

    def convert_pdf_to_docx(self, input_pdf, output_docx):
        cv = Converter(input_pdf)
        cv.convert(output_docx, start=0, end=None)
        cv.close()

    def convert_docx_to_pdf(self, input_docx, output_pdf):
        convert(input_docx, output_pdf)

if __name__ == "__main__":
    root = tk.Tk()
    app = DaveApp(root)
    root.mainloop()
