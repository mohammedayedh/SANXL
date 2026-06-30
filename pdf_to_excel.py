import customtkinter as ctk
import tkinter.filedialog as fd
import tkinter.messagebox as messagebox
from PIL import Image
import pdfplumber
import pandas as pd
import os
import threading
import subprocess
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class PDFtoExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window setup
        self.title("SANXL")
        self.geometry("600x550")
        self.resizable(False, False)
        
        # Appearance
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Logo and Title
        try:
            self.logo_image = ctk.CTkImage(light_image=Image.open(resource_path("logo.jpg")), dark_image=Image.open(resource_path("logo.jpg")), size=(80, 80))
            self.logo_label = ctk.CTkLabel(self, image=self.logo_image, text="")
            self.logo_label.pack(pady=(20, 0))
        except Exception as e:
            print("Logo not found:", e)

        self.title_label = ctk.CTkLabel(self, text="SANXL", font=ctk.CTkFont(size=28, weight="bold"))
        self.title_label.pack(pady=(5, 5))

        self.subtitle_label = ctk.CTkLabel(self, text="Extract tables from PDF files and save them to Excel", font=ctk.CTkFont(size=14), text_color="gray")
        self.subtitle_label.pack(pady=(0, 20))

        # Main Frame
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=40)

        # Input PDF Section
        self.input_label = ctk.CTkLabel(self.main_frame, text="Select PDF File:", anchor="w")
        self.input_label.grid(row=0, column=0, sticky="w", pady=(10, 5))

        self.input_entry = ctk.CTkEntry(self.main_frame, width=350, placeholder_text="Path to PDF file...")
        self.input_entry.grid(row=1, column=0, pady=5, padx=(0, 10))

        self.input_btn = ctk.CTkButton(self.main_frame, text="Browse", width=100, command=self.browse_input)
        self.input_btn.grid(row=1, column=1, pady=5)

        # Output Excel Section
        self.output_label = ctk.CTkLabel(self.main_frame, text="Save Excel File As:", anchor="w")
        self.output_label.grid(row=2, column=0, sticky="w", pady=(15, 5))

        self.output_entry = ctk.CTkEntry(self.main_frame, width=350, placeholder_text="Path to save Excel file...")
        self.output_entry.grid(row=3, column=0, pady=5, padx=(0, 10))

        self.output_btn = ctk.CTkButton(self.main_frame, text="Browse", width=100, command=self.browse_output)
        self.output_btn.grid(row=3, column=1, pady=5)

        # Options Section
        self.options_label = ctk.CTkLabel(self.main_frame, text="Extraction Strategy:", anchor="w")
        self.options_label.grid(row=4, column=0, sticky="w", pady=(15, 5))

        self.strategy_var = ctk.StringVar(value="Automatic")
        self.strategy_menu = ctk.CTkOptionMenu(self.main_frame, values=["Automatic", "Lines (Grid)", "Text (No Grid)"], variable=self.strategy_var, width=200)
        self.strategy_menu.grid(row=5, column=0, sticky="w", pady=5)

        # Convert Button
        self.convert_btn = ctk.CTkButton(self, text="Convert to Excel", font=ctk.CTkFont(size=16, weight="bold"), height=45, command=self.start_conversion)
        self.convert_btn.pack(pady=(20, 10))

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.pack(pady=(0, 10))
        self.progress_bar.set(0)

        # Status Label
        self.status_label = ctk.CTkLabel(self, text="", text_color="gray")
        self.status_label.pack(pady=(0, 15))

    def browse_input(self):
        filename = fd.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.input_entry.delete(0, 'end')
            self.input_entry.insert(0, filename)
            
            # Automatically set output filename if it's empty
            if not self.output_entry.get():
                base, _ = os.path.splitext(filename)
                self.output_entry.insert(0, base + "_output.xlsx")

    def browse_output(self):
        filename = fd.asksaveasfilename(
            title="Save Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_entry.delete(0, 'end')
            self.output_entry.insert(0, filename)

    def update_status(self, message, color="gray"):
        self.status_label.configure(text=message, text_color=color)
        self.update_idletasks()

    def update_progress(self, value):
        self.progress_bar.set(value)
        self.update_idletasks()

    def show_success(self, excel_path):
        if messagebox.askyesno("Success", f"Conversion completed successfully!\n\nFile saved at:\n{excel_path}\n\nDo you want to open the output folder?"):
            try:
                subprocess.run(["open", "-R", excel_path])
            except Exception as e:
                print("Error opening folder:", e)

    def show_warning(self):
        messagebox.showwarning("Warning", "No tables found in the PDF.")

    def show_error(self, err_msg):
        messagebox.showerror("Error", f"An error occurred:\n{err_msg}")

    def start_conversion(self):
        pdf_path = self.input_entry.get().strip()
        excel_path = self.output_entry.get().strip()
        strategy = self.strategy_var.get()

        if not pdf_path or not os.path.exists(pdf_path):
            self.update_status("Error: Please select a valid input PDF file.", "red")
            messagebox.showerror("Error", "Please select a valid input PDF file.")
            return
        
        if not excel_path:
            self.update_status("Error: Please specify the output Excel file.", "red")
            messagebox.showerror("Error", "Please specify the output Excel file.")
            return

        # Disable button during conversion
        self.convert_btn.configure(state="disabled", text="Converting...")
        self.update_status("Starting conversion...", "blue")
        self.update_progress(0)

        # Run conversion in a separate thread to keep GUI responsive
        threading.Thread(target=self.convert_logic, args=(pdf_path, excel_path, strategy), daemon=True).start()

    def convert_logic(self, pdf_path, excel_path, strategy):
        try:
            all_tables = []
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    self.update_status(f"Processing page {i+1} of {total_pages}...", "blue")
                    self.update_progress((i) / total_pages)
                    
                    table = None
                    if strategy in ["Automatic", "Lines (Grid)"]:
                        table = page.extract_table(table_settings={
                            "vertical_strategy": "lines", 
                            "horizontal_strategy": "lines",
                            "snap_tolerance": 3,
                        })
                    
                    if not table and strategy in ["Automatic", "Text (No Grid)"]:
                        table = page.extract_table(table_settings={
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text",
                        })
                    
                    if table:
                        # table is a list of lists. table[0] usually contains headers.
                        # Need to handle cases where table might be empty or invalid
                        if len(table) > 1:
                            df = pd.DataFrame(table[1:], columns=table[0])
                        else:
                            df = pd.DataFrame(table)
                        all_tables.append(df)
            
            if all_tables:
                self.update_status("Saving to Excel...", "blue")
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    for i, df in enumerate(all_tables):
                        df.to_excel(writer, sheet_name=f'Page_{i+1}', index=False)
                
                self.update_status(f"Success! Saved to {os.path.basename(excel_path)}", "green")
                self.update_progress(1.0)
                
                # Show completion notification safely in main thread
                self.after(0, self.show_success, excel_path)
            else:
                self.update_status("Warning: No tables found in the PDF.", "orange")
                self.update_progress(0)
                self.after(0, self.show_warning)
                
        except Exception as e:
            self.update_status(f"Error: {str(e)}", "red")
            self.after(0, self.show_error, str(e))
            
        finally:
            self.after(0, lambda: self.convert_btn.configure(state="normal", text="Convert to Excel"))

if __name__ == "__main__":
    app = PDFtoExcelApp()
    app.mainloop()
