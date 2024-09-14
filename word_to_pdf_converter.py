import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Function to convert Word to PDF
def convert_word_to_pdf(word_path, pdf_path):
    try:
        doc = Document(word_path)
        c = canvas.Canvas(pdf_path, pagesize=letter)
        width, height = letter

        text_object = c.beginText(60, height - 60)
        text_object.setFont("Open Sans", 16)
        
        for para in doc.paragraphs:
            text_object.textLines(para.text)
        
        c.drawText(text_object)
        c.showPage()
        c.save()
        messagebox.showinfo("Success", "Conversion completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Custom dialog for file selection
class CustomFileDialog:
    def __init__(self, parent, title, initialdir, filetypes, action):
        self.parent = parent
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("400x200")
        self.top.configure(bg="lightgray")

        tk.Label(self.top, text="File Name:", bg="lightgray", font=("Helvetica", 12)).pack(pady=10)
        
        self.file_entry = tk.Entry(self.top, width=40, font=("Helvetica", 12))
        self.file_entry.pack(pady=5)
        
        self.filetypes = filetypes
        self.action = action
        
        tk.Button(self.top, text="Browse", command=self.browse, bg="lightblue", fg="black", font=("Helvetica", 12)).pack(pady=5)
        tk.Button(self.top, text="OK", command=self.ok, bg="lightgreen", fg="black", font=("Helvetica", 12)).pack(pady=10)
        tk.Button(self.top, text="Cancel", command=self.cancel, bg="lightcoral", fg="black", font=("Helvetica", 12)).pack(pady=5)

# Function to open file dialog and convert the selected file
def browse_and_convert():
    word_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if not word_path:
        return

    pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return

    convert_word_to_pdf(word_path, pdf_path)

# Setup Tkinter GUI
root = tk.Tk()
root.title("Word to PDF Converter")
root.geometry("800x400")

tk.Button(root, text="Convert Word to PDF", command=browse_and_convert, font=("Helvetica", 14), bg="lightblue", fg="black").pack(pady=30)


tk.Label(root, text="Welcome to Word to PDF Converter", font=("Helvetica", 16), bg="lightgray", fg="darkblue").pack(pady=10)


# tk.Button(root, text="Convert Word to PDF", command=browse_and_convert).pack(pady=20)

root.mainloop()
