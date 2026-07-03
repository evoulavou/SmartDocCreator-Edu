import tkinter as tk
from tkinter import filedialog
from placement import placements
from notification3 import notifications3
from sort import organize_files_by_school

root = tk.Tk()
root.withdraw()

filename = filedialog.askopenfilename(
    title="Επίλεξε αρχείο",
    filetypes=[
        ("Excel", "*.xlsx")
    ]
)

print(filename)
placements(filename)
notifications3(filename)
organize_files_by_school('notifications/', 'my_school_pdfs/')