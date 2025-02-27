import os
import win32com.client
from pathlib import Path
import sys

def replace_bookmarks_in_word(file_name, replacements, start_number):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run Word in the background

    # Get the file path of the current working directory
    if getattr(sys, 'frozen', False):  
        script_dir = os.path.dirname(sys.executable)  # Voor .exe
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))  # Voor script

    file_path = os.path.join(script_dir, file_name)  
    doc = word.Documents.Open(file_path)

    modified_bookmarks = []  # Keep track of modified bookmarks

    for i, replacement in enumerate(replacements):
        bookmark_name = f"n{start_number + i}"  # Start from the chosen number

        if doc.Bookmarks.Exists(bookmark_name):
            bookmark = doc.Bookmarks(bookmark_name)
            bookmark_range = bookmark.Range  

            start, end = bookmark_range.Start, bookmark_range.End  
            bookmark_range.Text = replacement  

            # Restore the bookmark (since Word deletes it)
            doc.Bookmarks.Add(bookmark_name, doc.Range(start, start + len(replacement)))

            modified_bookmarks.append(bookmark_name)  # Track modified bookmark
            print(f"✅ Bookmark nummer '{bookmark_name}' verandert met '{replacement}'")
        else:
            print(f"⚠️ Error: Bookmark nummer '{bookmark_name}' bestaat niet, probeer opnieuw! (Kies een lager nummer om mee te beginnen of voer in minder sticker nummers!)")
            exit()

    # Now delete all unmodified bookmarks and their text
    for i in range(1, 326):  # From n1 to n325
        bookmark_name = f"n{i}"
        if bookmark_name not in modified_bookmarks:  # If this bookmark wasn't modified
            if doc.Bookmarks.Exists(bookmark_name):
                bookmark = doc.Bookmarks(bookmark_name)
                # Clear the text inside the bookmark
                bookmark_range = bookmark.Range
                bookmark_range.Delete()  # This deletes the text inside the bookmark

                # Check if bookmark still exists before deleting
                if doc.Bookmarks.Exists(bookmark_name):
                    doc.Bookmarks.Item(bookmark_name).Delete()

    # Save to the Downloads folder as .doc format only
    downloads_folder = str(Path.home() / "Downloads")  # Get the Downloads folder
    output_path_doc = os.path.join(downloads_folder, "Word_template_USB_nummers.doc")  # Save as .doc format

    # Save it as a .doc file
    doc.SaveAs(output_path_doc, FileFormat=0)  # FileFormat=0 saves the file as .doc format
    
    # Close the document and quit Word
    doc.Close()
    word.Quit()

    print(f"✅ Word document is succesvol gemaakt! Opgeslagen in: {output_path_doc} \n🖥️  Je kan nu de terminal sluiten ✔✔✔")

# --- Get user input ---
print("❌ Laat de terminal open staan (anders gaat het niet werken) ❌")
start_number = input("Vanaf welk nummer wil je beginnen?: ").strip()
while not start_number.isdigit():
    start_number = input("Voer een geldig getal in: ").strip()
start_number = int(start_number)  # Convert to integer

print("Sticker nummers (klik enter na elk nummer, double enter om te stoppen):")
numbers = []
while True:
    line = input().strip()
    if line == "":  # Stop when user presses Enter on an empty line
        break
    numbers.append(line)

# Run the function
replace_bookmarks_in_word("template.docx", numbers, start_number)
input("Druk op Enter om af te sluiten.")
