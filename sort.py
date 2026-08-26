import os
import re
import shutil
from collections import defaultdict
from docx import Document


def sanitize_foldername(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()


def read_tags_from_docx(file_path):
    try:
        doc = Document(file_path)
        props = doc.core_properties

        # The 'keywords' property maps to 'Ετικέτες' in Word
        raw_tags = props.keywords

        if raw_tags:
            schools = [s.strip() for s in raw_tags.split(',') if s.strip()]
            return schools
        else:
            return []

    except Exception as e:
        print(f"  Error reading '{os.path.basename(file_path)}': {e}")
        return []
        
def prepare_pdfs_for_signing(input_folder, output_folder):
    """
    Αντιγράφει όλα τα ανυπόγραφα PDF σε έναν καθαρό φάκελο
    για μαζική ψηφιακή υπογραφή.
    """
    os.makedirs(output_folder, exist_ok=True)

    copied = 0

    for filename in os.listdir(input_folder):
        if not filename.lower().endswith(".pdf"):
            continue

        # Δεν θέλουμε τυχόν ήδη υπογεγραμμένα αρχεία
        if filename.lower().endswith("_signed.pdf"):
            continue

        source = os.path.join(input_folder, filename)
        destination = os.path.join(output_folder, filename)

        shutil.copy2(source, destination)
        copied += 1

    print(f"Έτοιμα για υπογραφή: {copied} PDF")
    print(f"Φάκελος: {output_folder}")

def organize_signed_pdfs_by_school(metadata_folder, signed_folder, output_folder):
    """
    Ταξινομεί τα υπογεγραμμένα *_signed.pdf ανά σχολείο,
    χρησιμοποιώντας τις ετικέτες του αντίστοιχου αρχικού .docx.
    """

    os.makedirs(output_folder, exist_ok=True)

    total_signed = 0
    matched = 0
    unmatched = []

    for filename in os.listdir(signed_folder):

        # Μας ενδιαφέρουν μόνο τα υπογεγραμμένα PDF
        if not filename.lower().endswith("_signed.pdf"):
            continue

        total_signed += 1

        signed_pdf = os.path.join(signed_folder, filename)

        # Αφαιρούμε το _signed.pdf
        name_without_signed = filename[:-len("_signed.pdf")]

        # Αφαιρούμε και το τελευταίο τμήμα χαρακτήρων
        # που προσθέτει το πρόγραμμα ψηφιακής υπογραφής
        base_name = name_without_signed.rsplit("_", 1)[0]

        # Βρίσκουμε το αντίστοιχο αρχικό Word.
# Το πρόγραμμα ψηφιακής υπογραφής μετατρέπει τις τελείες
# του αρχικού ονόματος σε κάτω παύλες.

docx_path = None

normalized_base_name = base_name.replace(".", "_")

for original_filename in os.listdir(metadata_folder):

    if not original_filename.lower().endswith(".docx"):
        continue

    original_base_name = os.path.splitext(original_filename)[0]
    normalized_original = original_base_name.replace(".", "_")

    if normalized_original == normalized_base_name:
        docx_path = os.path.join(metadata_folder, original_filename)
        break

if docx_path is None:
    print(f"ΧΩΡΙΣ ΑΝΤΙΣΤΟΙΧΙΣΗ: {filename}")
    unmatched.append(filename)
    continue

        # Διαβάζουμε τα σχολεία από τις ετικέτες του Word
schools = read_tags_from_docx(docx_path)

if not schools:
    print(f"ΧΩΡΙΣ ΣΧΟΛΕΙΟ: {filename}")
    unmatched.append(filename)
    continue

matched += 1

        # Το ίδιο κοινοποιητήριο μπορεί να αφορά περισσότερα σχολεία
        for school in schools:

            school_folder = os.path.join(
                output_folder,
                sanitize_foldername(school)
            )

            os.makedirs(school_folder, exist_ok=True)

            destination = os.path.join(
                school_folder,
                filename
            )

            shutil.copy2(signed_pdf, destination)

    print()
    print("===== ΑΠΟΤΕΛΕΣΜΑ ΤΑΞΙΝΟΜΗΣΗΣ =====")
    print(f"Υπογεγραμμένα PDF: {total_signed}")
    print(f"Αντιστοιχίστηκαν: {matched}")
    print(f"Χωρίς αντιστοίχιση: {len(unmatched)}")

    if unmatched:
        print()
        print("Αρχεία που χρειάζονται έλεγχο:")
        for filename in unmatched:
            print(f" - {filename}")

def organize_files_by_school(input_folder, output_folder):
    if not os.path.exists(input_folder):
        print(f"Error: The folder '{input_folder}' does not exist.")
        return

    os.makedirs(output_folder, exist_ok=True)
    print(f"Scanning .docx files in '{input_folder}'...\n")

    # Maps school name -> list of (docx_path, pdf_path_or_None)
    school_map = defaultdict(list)

    for filename in os.listdir(input_folder):
        if not filename.lower().endswith(".docx"):
            continue

        docx_path = os.path.join(input_folder, filename)
        base_name = os.path.splitext(filename)[0]
        pdf_filename = base_name + ".pdf"
        pdf_path = os.path.join(input_folder, pdf_filename)

        print(f"Processing: {filename}")
        schools = read_tags_from_docx(docx_path)

        if not schools:
            print(f"  → No 'Ετικέτες' found. Skipping.\n")
            continue

        print(f"  → Schools found: {schools}")

        for school in schools:
            school_map[school].append({
                "docx_path": docx_path,
                "docx_filename": filename,
                "pdf_path": pdf_path if os.path.exists(pdf_path) else None,
                "pdf_filename": pdf_filename,
            })

        print()

    if not school_map:
        print("No files with 'Ετικέτες' were found.")
        return

    print(f"Found {len(school_map)} unique school(s). Organizing files...\n")

    for school_name, file_entries in school_map.items():
        safe_name = sanitize_foldername(school_name)
        school_folder = os.path.join(output_folder, safe_name)
        os.makedirs(school_folder, exist_ok=True)

        for entry in file_entries:
            dest_docx = os.path.join(school_folder, entry["docx_filename"])
            shutil.copy2(entry["docx_path"], dest_docx)
            print(f"  [{safe_name}] Copied: {entry['docx_filename']}")

            if entry["pdf_path"]:
                dest_pdf = os.path.join(school_folder, entry["pdf_filename"])
                shutil.copy2(entry["pdf_path"], dest_pdf)
                print(f"  [{safe_name}] Copied: {entry['pdf_filename']}")
            else:
                print(f"  [{safe_name}] Warning: PDF not found for '{entry['docx_filename']}'")

        print()

    print("Done! Files organized in:", output_folder)


if __name__ == "__main__":
    INPUT_FOLDER = "my_docx_files"
    OUTPUT_FOLDER = "my_schools_pdfs"

    organize_files_by_school(INPUT_FOLDER, OUTPUT_FOLDER)
