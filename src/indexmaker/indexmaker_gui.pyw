import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import os
from indexmaker.indexmaker_script import IndexMaker

# Function to select a PDF file
def select_pdf():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF Files", "*.pdf")],
        title="Select a PDF File"
    )
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

# Function to select an optional index file
def select_index_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("All Files", "*.*")],
        title="Select an Index File"
    )
    if file_path:
        index_entry.delete(0, tk.END)
        index_entry.insert(0, file_path)

# Function to run the IndexMaker
def run_script():
    pdf_file = pdf_entry.get()
    index_file = index_entry.get()
    output_file = output_entry.get()
    start_page = start_page_entry.get()

    if not pdf_file:
        messagebox.showerror("Error", "Please select a PDF file.")
        return

    try:
        # Convert start page to integer
        start_page = int(start_page) if start_page.strip() else 1

        # Create and run the IndexMaker
        index_maker = IndexMaker(
            input_file_path=pdf_file,
            output_docx=output_file,
            start_from_page=start_page,
            own_index=index_file if index_file.strip() else None
        )
        index_maker.main()

        # Display success message
        output_area.delete(1.0, tk.END)
        output_area.insert(tk.END, f"Index created successfully: {output_file}\n")

        # Enable the "Open Output File" button
        if os.path.exists(output_file):
            output_button.config(state=tk.NORMAL, bg="lightgreen")
        else:
            output_button.config(state=tk.DISABLED)

    except Exception as e:
        output_area.insert(tk.END, f"Error: {e}\n")

# Function to open the output file
def open_output():
    output_file = output_entry.get()
    if os.path.exists(output_file):
        os.startfile(output_file) if os.name == "nt" else os.system(f"xdg-open {output_file}")
    else:
        messagebox.showerror("Error", "Output file not found.")

# Create the GUI
root = tk.Tk()
root.title("IndexMaker GUI")

#intro label
explanations = """
IndexMaker
----------------------
(c) Nicolas Soler 2025
Enter a pdf file path, an optional index file path, 
output file path, and starting page number.

The script will generate an index of all
non-common words found in the PDF (in docx format).
"""
intro_label = tk.Label(root, text=explanations, font=("Helvetica", 12), bg="lightyellow", justify=tk.LEFT)
intro_label.pack(pady=10)

# PDF file selection
pdf_label = tk.Label(root, text="Select PDF File:")
pdf_label.pack(pady=5)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.pack(pady=5)
pdf_button = tk.Button(root, text="Browse", command=select_pdf)
pdf_button.pack(pady=5)

# Optional index file selection
index_label = tk.Label(root, text="Optional Index File:")
index_label.pack(pady=5)
index_entry = tk.Entry(root, width=50)
index_entry.pack(pady=5)
index_button = tk.Button(root, text="Browse", command=select_index_file)
index_button.pack(pady=5)

# Output file path
output_label = tk.Label(root, text="Output File Path:")
output_label.pack(pady=5)
output_entry = tk.Entry(root, width=50)
output_entry.insert(0, "output_index.docx")  # Default output file name
output_entry.pack(pady=5)

# Start page input
start_page_label = tk.Label(root, text="Page number of the document you consider as the first page:")
start_page_label.pack(pady=5)
start_page_entry = tk.Entry(root, width=50)
start_page_entry.insert(0, "1")  # Default start page
start_page_entry.pack(pady=5)

# Run button
run_button = tk.Button(root, text="Run!", bg="blue", fg="white", command=run_script)
run_button.pack(pady=10)

# Output area
output_label = tk.Label(root, text="Script Output:")
output_label.pack(pady=5)
output_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=10)
output_area.pack(pady=5)

# Output file button
output_button = tk.Button(root, text="Open Output File", command=open_output, state=tk.DISABLED)
output_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
