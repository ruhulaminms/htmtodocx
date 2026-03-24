import os
import win32com.client as win32

# The path of the folder where your HTML files are located (e.g.: D:/Projects/Files)
# Here, the current folder is being used
path = os.getcwd()
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False  # The Word window will not appear while processing

# Checking all files in the folder
for filename in os.listdir(path):
    if filename.endswith(".htm") or filename.endswith(".html"):
        # Create the full path of the file
        file_abs_path = os.path.join(path, filename)
        output_path = os.path.join(path, os.path.splitext(filename)[0] + ".docx")
        
        print(f"Processing: {filename}...")
        
        try:
            # Open the HTML file using Word
            doc = word.Documents.Open(file_abs_path)
            # Save in docx format (FileFormat=16 means docx)
            doc.SaveAs(output_path, FileFormat=16)
            doc.Close()
        except Exception as e:
            print(f"Error occurred in {filename}: {e}")

word.Quit()
print("\nCongratulations! All files have been successfully saved in .docx format.")