import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

def copy_pdfs_flat(source_folder, target_folder):
    count = 0
    for root, dirs, files in os.walk(source_folder):
        for file in files:
            if file.lower().endswith('.pdf'):
                source_path = os.path.join(root, file)
                destination_path = os.path.join(target_folder, file)

                # Handle duplicate file names
                base, ext = os.path.splitext(file)
                counter = 1
                while os.path.exists(destination_path):
                    new_name = f"{base}_{counter}{ext}"
                    destination_path = os.path.join(target_folder, new_name)
                    counter += 1

                shutil.copy2(source_path, destination_path)
                count += 1

    return count

def browse_source():
    folder = filedialog.askdirectory()
    if folder:
        source_var.set(folder)

def browse_target():
    folder = filedialog.askdirectory()
    if folder:
        target_var.set(folder)

def start_copy():
    source = source_var.get()
    target = target_var.get()

    if not source or not target:
        messagebox.showwarning("Missing Input", "Please select both source and target folders.")
        return

    copied = copy_pdfs_flat(source, target)
    messagebox.showinfo("Done", f"✅ Total PDF files copied: {copied}")

# === GUI ===
root = tk.Tk()
root.title("PDF Copier (Flat Copy)")
root.geometry("500x220")
root.resizable(False, False)

source_var = tk.StringVar()
target_var = tk.StringVar()

tk.Label(root, text="Select Source Folder:").pack(pady=5)
tk.Entry(root, textvariable=source_var, width=60).pack()
tk.Button(root, text="Browse", command=browse_source).pack(pady=5)

tk.Label(root, text="Select Target Folder:").pack(pady=5)
tk.Entry(root, textvariable=target_var, width=60).pack()
tk.Button(root, text="Browse", command=browse_target).pack(pady=5)

tk.Button(root, text="Copy PDFs", command=start_copy, bg="#4CAF50", fg="white", padx=20, pady=5).pack(pady=15)

root.mainloop()
