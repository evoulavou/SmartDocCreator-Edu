import tkinter as tk
from tkinter import filedialog, messagebox

from placement import placements
from notification3 import notifications3
from sort import (
    organize_files_by_school,
    prepare_pdfs_for_signing,
    organize_signed_pdfs_by_school
)


def create_notifications():
    filename = filedialog.askopenfilename(
        title="Επίλεξε αρχείο Excel",
        filetypes=[
            ("Excel", "*.xlsx")
        ]
    )

    if not filename:
        return

    print(filename)

    placements(filename)
    notifications3(filename)

    prepare_pdfs_for_signing(
        "notifications/",
        "to_sign/"
    )

    messagebox.showinfo(
        "Ολοκληρώθηκε",
        "Τα κοινοποιητήρια δημιουργήθηκαν.\n\n"
        "Τα PDF για μαζική υπογραφή βρίσκονται "
        "στον φάκελο 'to_sign'."
    )


def sort_signed_notifications():
    signed_folder = filedialog.askdirectory(
        title="Επίλεξε τον φάκελο με τα υπογεγραμμένα PDF"
    )

    if not signed_folder:
        return

    organize_signed_pdfs_by_school(
        "notifications/",
        signed_folder,
        "my_school_pdfs/"
    )

    messagebox.showinfo(
        "Ολοκληρώθηκε",
        "Τα υπογεγραμμένα κοινοποιητήρια "
        "ταξινομήθηκαν ανά σχολείο."
    )


root = tk.Tk()

root.title("SmartDocCreator - ΔΔΕ Χαλκιδικής")
root.geometry("520x240")

title_label = tk.Label(
    root,
    text="SmartDocCreator",
    font=("Arial", 16, "bold")
)

title_label.pack(pady=20)


button_create = tk.Button(
    root,
    text="1. Δημιουργία κοινοποιητηρίων για υπογραφή",
    command=create_notifications,
    width=48,
    height=3
)

button_create.pack(pady=10)


button_sort = tk.Button(
    root,
    text="2. Ταξινόμηση υπογεγραμμένων ανά σχολείο",
    command=sort_signed_notifications,
    width=48,
    height=3
)

button_sort.pack(pady=10)


root.mainloop()
