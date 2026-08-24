import tkinter as tk
from tkinter import filedialog
from placement import placements
from notification3 import notifications3
from sort import organize_files_by_school, prepare_pdfs_for_signing

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
prepare_pdfs_for_signing('notifications/', 'to_sign/')
