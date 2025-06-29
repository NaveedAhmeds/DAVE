import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import os
from PIL import Image, ImageTk
from pdf2docx import Converter
from docx2pdf import convert

# loading screen logic with typing effect and 1-second delay
def show_splash(root):
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    splash.geometry("400x300+500+250")
    splash.configure(bg="#f4eee8")

    # Loading and showing the image...
    try:
        logo_img = Image.open("logo.jpg")
        logo_img = logo_img.resize((120, 120), Image.Resampling.LANCZOS)
        logo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(splash, image=logo, bg="#f4eee8")
        logo_label.image = logo  # Keep a reference
        logo_label.pack(pady=(30, 10))
    except Exception as e:
        print("Logo load error:", e)

    # Typing title label effect...
    title_label = tk.Label(
        splash,
        text="",
        font=("Georgia", 28, "bold italic"),
        fg="#5b4636",
        bg="#f4eee8"
    )
    title_label.pack()

    # Subtitle for the intro...
    subtitle = tk.Label(
        splash,
        text="Initializing... Converting the impossible.",
        font=("Helvetica", 12),
        fg="#5b4636",
        bg="#f4eee8",
        pady=20
    )
    subtitle.pack()

    # *** My Credit Line...  (hehehehhehehe)***
    credit = tk.Label(
        splash,
        text="© 2025 by Naveed -Dev.",
        font=("Helvetica", 10, "italic"),
        fg="#7a6a57",
        bg="#f4eee8",
        pady=5
    )
    credit.pack()

    splash_text = "D.A.V.E"

    def type_text(index=0):
        if index <= len(splash_text):
            title_label.config(text=splash_text[:index])
            splash.after(150, type_text, index + 1)
        else:
            splash.after(1000, lambda: [splash.destroy(), root.deiconify()])

    root.withdraw()
    type_text()


#The app itself (GUI)...
class DaveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("D.A.V.E. - Document Adapter & Versatile Encoder")
        self.root.geometry("800x600")

        #Note: (if you are in here hello..! please don't modify this as it will ruin the background, thank you.)
        self.root.resizable(False, False)

        #The window icon...
        try:
            self.logo = tk.PhotoImage(file="logo.png")
            self.root.iconphoto(False, self.logo)
        except Exception as e:
            print("Could not load app icon:", e)

        # Background Renissance wooohoo...
        self.bg_image = Image.open("bg.jpeg")
        self.bg_image = self.bg_image.resize((800, 600), Image.Resampling.LANCZOS)
        self.bg_photo = ImageTk.PhotoImage(self.bg_image)

        self.background_label = tk.Label(root, image=self.bg_photo)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

        # Title label...
        self.title_label = tk.Label(
            root,
            text="D.A.V.E",
            font=("Georgia", 48, "bold italic"),
            bg="#f4eee8",
            fg="#5b4636"
        )
        self.title_label.place(relx=0.5, rely=0.12, anchor="center")

        # File label...
        self.file_label = tk.Label(
            root,
            text="",
            font=("Helvetica", 13, "bold"),
            bg="#f4eee8",
            fg="#5b4636"
        )
        self.file_label.place(relx=0.5, rely=0.4, anchor="center")

        # Buttons... (don't press me buttons hahahaha)
        # This project did press me buttons...
        self.buttons_frame = tk.Frame(root, bg="#f4eee8")
        self.buttons_frame.place(relx=0.5, rely=0.3, anchor="center")

        self.btn_pdf_to_docx = ttk.Button(
            self.buttons_frame,
            text="Convert PDF -> DOCX",
            command=self.select_pdf_file,
            width=25
        )
        self.btn_pdf_to_docx.pack(side=tk.LEFT, padx=20)

        self.btn_docx_to_pdf = ttk.Button(
            self.buttons_frame,
            text="Convert DOCX -> PDF",
            command=self.select_docx_file,
            width=25
        )
        self.btn_docx_to_pdf.pack(side=tk.LEFT, padx=20)

        # Progress bar setup...
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

        # I am hiding it here while theres no file...
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

        # Quote at bottom... (Because I am addicted to philosophy)
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

        # Status bar...
        self.status_var = tk.StringVar()
        self.status_var.set("Welcome to D.A.V.E! Select a conversion to start.")
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

    # Opens file dialog to select PDF, updates UI and starts conversion...
    def select_pdf_file(self):
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf")],
        )
        if file_path:
            self.status_var.set(f"Selected PDF: {file_path}")
            self.file_label.config(text=f"Converting: {os.path.basename(file_path)}")
            self.start_conversion(file_path, "pdf_to_docx")

    # Opens file dialog to select DOCX, updates UI and starts conversion...
    def select_docx_file(self):
        file_path = filedialog.askopenfilename(
            title="Select DOCX File",
            filetypes=[("Word Documents", "*.docx")],
        )
        if file_path:
            self.status_var.set(f"Selected DOCX: {file_path}")
            self.file_label.config(text=f"Converting: {os.path.basename(file_path)}")
            self.start_conversion(file_path, "docx_to_pdf")

    # Disables buttons, shows progress bar, and launches conversion thread...
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

    # Runs conversion, progress, calls convert functions...
    def run_conversion(self, file_path, conversion_type):
        try:
            base, ext = os.path.splitext(file_path)
            output_file = base + ("_daved.docx" if conversion_type == "pdf_to_docx" else "_daved.pdf")

            # Progress Bar
            for i in range(101):
                time.sleep(0.003)
                self.root.after(0, self.update_progress, i)

            # conversion call after progress bar done...
            if conversion_type == "pdf_to_docx":
                self.convert_pdf_to_docx(file_path, output_file)
            else:
                self.convert_docx_to_pdf(file_path, output_file)

        except Exception as e:
            self.root.after(0, self.handle_error, str(e))
            return

        # Notify main thread to finish up...
        self.root.after(0, self.finish_conversion, output_file)

    # Update progress bar and label...
    def update_progress(self, value):
        self.progress['value'] = value
        self.progress_label.config(text=f"Progress: {value}%")

    # Finish conversion, asking to save file, reset UI...
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

    # Error handling: resetting UI and show error popup...
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

    # Use pdf2docx to convert PDF to DOCX
    # Note: Dev take: make sure you install the packages following the readme closely, 
    # The application relies on these frameworks for the conversion
    def convert_pdf_to_docx(self, input_pdf, output_docx):
        cv = Converter(input_pdf)
        cv.convert(output_docx, start=0, end=None)
        cv.close()

    # Use docx2pdf to convert DOCX to PDF
    def convert_docx_to_pdf(self, input_docx, output_pdf):
        convert(input_docx, output_pdf)


# Launching the main app...
if __name__ == "__main__":
    root = tk.Tk()
    show_splash(root)
    app = DaveApp(root)
    root.mainloop()
