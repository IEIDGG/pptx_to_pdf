import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
from tkinterdnd2 import DND_FILES, TkinterDnD

class PPTXtoPDFConverter:
    def __init__(self, root: TkinterDnD.Tk):
        self.root = root
        self.root.title("PPTX to PDF Converter")
        self.root.geometry("600x400")

        self.pptx_files = []
        self.output_directory = ""

        self.file_list_frame = tk.Frame(self.root)
        self.file_list_frame.pack(pady=10, padx=10, fill="x")

        self.lbl_files = tk.Label(self.file_list_frame, text="PPTX Files:")
        self.lbl_files.pack(side="left")

        self.listbox_files = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=70, height=10)
        self.listbox_files.pack(pady=5, padx=10, fill="x")
        
        self.listbox_files.drop_target_register(DND_FILES)
        self.listbox_files.dnd_bind('<<Drop>>', self.drop_files)
        
        self.buttons_frame = tk.Frame(self.root)
        self.buttons_frame.pack(pady=10, padx=10, fill="x")

        self.btn_add_files = tk.Button(self.buttons_frame, text="Add PPTX Files", command=self.add_files)
        self.btn_add_files.pack(side="left", padx=5)

        self.btn_remove_files = tk.Button(self.buttons_frame, text="Remove Selected", command=self.remove_files)
        self.btn_remove_files.pack(side="left", padx=5)

        self.output_dir_frame = tk.Frame(self.root)
        self.output_dir_frame.pack(pady=10, padx=10, fill="x")

        self.btn_select_output_dir = tk.Button(self.output_dir_frame, text="Select Output Directory", command=self.select_output_directory)
        self.btn_select_output_dir.pack(side="left", padx=5)

        self.lbl_output_dir_val = tk.Label(self.output_dir_frame, text="No directory selected", relief="sunken", width=40, anchor="w")
        self.lbl_output_dir_val.pack(side="left", padx=5, fill="x", expand=True)

        self.btn_convert = tk.Button(self.root, text="Convert to PDF", command=self.convert_to_pdf, state=tk.DISABLED)
        self.btn_convert.pack(pady=20, padx=10)

        self.lbl_status = tk.Label(self.root, text="", fg="green")
        self.lbl_status.pack(pady=5, padx=10)

        self.root.resizable(True, True)
        

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="Select PPTX Files",
            filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*"))
        )
        if files:
            for file_path in files:
                if file_path not in self.pptx_files and file_path.lower().endswith(".pptx"):
                    self.pptx_files.append(file_path)
                    self.listbox_files.insert(tk.END, os.path.basename(file_path))
            if self.pptx_files: # Check if any valid files were actually added
                self.update_convert_button_state()
                self.lbl_status.config(text=f"{len(self.listbox_files.get(0, tk.END))} file(s) in list.")
            else:
                self.lbl_status.config(text="No valid .pptx files selected or added.", fg="orange")

    def drop_files(self, event):
        raw_file_paths = event.data
        if '{' in raw_file_paths:
            import re
            file_paths = re.findall(r'{(.*?)}', raw_file_paths)
            if not file_paths:
                 temp_paths = []
                 current_path = ""
                 in_brace = False
                 for char in raw_file_paths:
                     if char == '{':
                         in_brace = True
                         if current_path.strip():
                             temp_paths.append(current_path.strip())
                             current_path = ""
                     elif char == '}':
                         in_brace = False
                         if current_path.strip():
                             temp_paths.append(current_path.strip())
                             current_path = ""
                     elif in_brace:
                         current_path += char
                     elif char == ' ' and not in_brace and current_path.strip():
                         temp_paths.append(current_path.strip())
                         current_path = ""
                     elif char != ' ':
                         current_path += char
                 if current_path.strip():
                     temp_paths.append(current_path.strip())
                 file_paths = temp_paths
        else:
            file_paths = raw_file_paths.split()

        added_count = 0
        for file_path in file_paths:
            file_path = file_path.strip('"') 
            if os.path.isfile(file_path) and file_path.lower().endswith(".pptx"):
                if file_path not in self.pptx_files:
                    self.pptx_files.append(file_path)
                    self.listbox_files.insert(tk.END, os.path.basename(file_path))
                    added_count += 1
            else:
                print(f"Skipping non-pptx or non-file: {file_path}") # For debugging
        
        if added_count > 0:
            self.update_convert_button_state()
            self.lbl_status.config(text=f"{added_count} file(s) dropped and added.", fg="green")
        elif file_paths:
            self.lbl_status.config(text="No valid .pptx files found in dropped items.", fg="orange")

    def remove_files(self):
        selected_indices = self.listbox_files.curselection()
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select files to remove.")
            return

        for index in sorted(selected_indices, reverse=True):
            self.listbox_files.delete(index)
            del self.pptx_files[index]
        
        self.update_convert_button_state()
        self.lbl_status.config(text=f"{len(selected_indices)} file(s) removed.")

    def select_output_directory(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_directory = directory
            self.lbl_output_dir_val.config(text=self.output_directory)
            self.update_convert_button_state()
            self.lbl_status.config(text=f"Output directory: {self.output_directory}")

    def update_convert_button_state(self):
        if self.pptx_files and self.output_directory:
            self.btn_convert.config(state=tk.NORMAL)
        else:
            self.btn_convert.config(state=tk.DISABLED)

    def convert_to_pdf(self):
        if not self.pptx_files:
            messagebox.showerror("Error", "No PPTX files selected.")
            return
        if not self.output_directory:
            messagebox.showerror("Error", "No output directory selected.")
            return

        self.lbl_status.config(text="Converting...", fg="blue")
        self.root.update_idletasks()

        soffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe" 

        converted_count = 0
        error_count = 0

        try:
            subprocess.run([soffice_path, "--version"], capture_output=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            error_msg = f"LibreOffice command '{soffice_path}' not found or not working. Please ensure LibreOffice is installed and '{soffice_path}' is in your system PATH or provide the full path in the script.\nError: {e}"
            messagebox.showerror("LibreOffice Error", error_msg)
            self.lbl_status.config(text="LibreOffice not found or error.", fg="red")
            return

        for pptx_file_path in self.pptx_files:
            try:
                file_name = os.path.basename(pptx_file_path)
                os.makedirs(self.output_directory, exist_ok=True)

                self.lbl_status.config(text=f"Converting: {file_name}...", fg="blue")
                self.root.update_idletasks()
                
                pdf_file_name = os.path.splitext(file_name)[0] + ".pdf"
                expected_pdf_path = os.path.join(self.output_directory, pdf_file_name)

                if os.path.exists(expected_pdf_path):
                    if not messagebox.askyesno("File Exists", f"{pdf_file_name} already exists in the output directory. Overwrite?"):
                        self.lbl_status.config(text=f"Skipped (overwrite): {file_name}", fg="orange")
                        self.root.update_idletasks()
                        continue

                cmd = [
                    soffice_path,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", self.output_directory,
                    pptx_file_path
                ]
                
                creation_flags = 0
                if os.name == 'nt':
                    creation_flags = subprocess.CREATE_NO_WINDOW

                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=creation_flags)
                stdout, stderr = process.communicate()

                if process.returncode == 0:
                    if os.path.exists(expected_pdf_path):
                        converted_count += 1
                        self.lbl_status.config(text=f"Converted: {file_name}", fg="green")
                    else:
                        error_count += 1
                        err_details = stderr.decode(errors='ignore').strip()
                        if not err_details:
                            err_details = stdout.decode(errors='ignore').strip()
                        error_msg_detail = f"Conversion of {file_name} reported success, but output PDF not found at {expected_pdf_path}. Details: {err_details if err_details else 'No details from LibreOffice.'}"
                        print(error_msg_detail)
                        if error_count <= 2:
                             messagebox.showerror("Conversion Output Error", error_msg_detail)
                        self.lbl_status.config(text=f"Output missing for: {file_name}", fg="red")
                else:
                    error_count += 1
                    err_output = stderr.decode(errors='ignore').strip()
                    if not err_output:
                        err_output = stdout.decode(errors='ignore').strip()
                    error_msg = f"Error converting {file_name}. LibreOffice exit code: {process.returncode}\nDetails: {err_output}"
                    print(error_msg)
                    if error_count <= 2:
                         messagebox.showerror("Conversion Error", error_msg)
                    self.lbl_status.config(text=f"Error converting {file_name}", fg="red")
                
                self.root.update_idletasks()

            except Exception as e:
                error_count += 1
                error_msg_script = f"Script error during conversion of {os.path.basename(pptx_file_path)}: {str(e)}"
                print(error_msg_script)
                if error_count <= 2:
                    messagebox.showerror("Conversion Script Error", error_msg_script)
                self.lbl_status.config(text=f"Script error for {os.path.basename(pptx_file_path)}", fg="red")
                self.root.update_idletasks()

        final_message = f"Conversion complete. {converted_count} file(s) converted successfully."
        if error_count > 0:
            final_message += f" {error_count} file(s) failed."
            messagebox.showwarning("Conversion Summary", final_message)
            self.lbl_status.config(text=final_message, fg="orange")
        else:
            messagebox.showinfo("Conversion Complete", final_message)
            self.lbl_status.config(text=final_message, fg="green")
            
        self.listbox_files.delete(0, tk.END)
        self.pptx_files.clear()
        self.update_convert_button_state()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PPTXtoPDFConverter(root)
    root.mainloop() 