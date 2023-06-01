import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
import win32com.client

# ユーザーインターフェイスの作成
class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.select_button = tk.Button(self)
        self.select_button["text"] = "エクセルファイルを選択"
        self.select_button["command"] = self.select_files
        self.select_button.pack(side="top")

        self.file_listbox = tk.Listbox(self, height=4, state='disabled')
        self.file_listbox.pack(side="top")

        self.convert_button = tk.Button(self)
        self.convert_button["text"] = "pdfファイルに変換"
        self.convert_button["command"] = self.convert_to_pdf
        self.convert_button.pack(side="top")

        self.progress = ttk.Progressbar(self, length=400, mode='determinate')
        self.progress.pack(side="top")

        self.status_text = tk.StringVar()
        self.status_label = tk.Label(self, textvariable=self.status_text)
        self.status_label.pack(side="top")

        self.quit = tk.Button(self, text="QUIT", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def select_files(self):
        self.filenames = filedialog.askopenfilenames(initialdir = "/",title = "Select files",filetypes = (("Excel files","*.xls*"),("all files","*.*")))
        print(f"Selected files: {self.filenames}")

        self.file_listbox['state'] = 'normal'
        self.file_listbox.delete(0, tk.END)
        for file in self.filenames:
            self.file_listbox.insert(tk.END, file)
        self.file_listbox['state'] = 'disabled'

    def convert_to_pdf(self):
        output_dir = filedialog.askdirectory()
        print(f"Output directory: {output_dir}")

        self.progress['maximum'] = len(self.filenames)
        self.progress['value'] = 0

        for filename in self.filenames:
            print(f"Converting {filename} to PDF...")
            self.status_text.set(f"Converting {filename} to PDF...")

            # Open the workbook
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False

            wb = excel.Workbooks.Open(filename)
            
            # Set page layout
            ws = wb.ActiveSheet
            ws.PageSetup.Orientation = 2  # 2 represents Landscape
            ws.PageSetup.PrintArea = "A1:AC52"

            # Export to PDF
            output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(filename))[0] + ".pdf")
            ws.ExportAsFixedFormat(0, output_file)
            wb.Close(True)
            excel.Quit()

            print(f"Conversion completed. Output file: {output_file}")
            self.progress['value'] += 1
            self.update_idletasks()

        self.status_text.set("Conversion completed.")

# メインの実行部分
root = tk.Tk()
app = Application(master=root)
app.mainloop()
