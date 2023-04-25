import os
import win32com.client as win32
print("01 or 60?")
n = input()
q = n
print("Type Qx 202x (ie Q2 2023)")
d = input()
# Specify the path to the folder containing .doc files
if (n == "01"):
    q = "1"
folder_path = "C:\\Users\\johnd\\Desktop\\" + d + "\\" + d + "\\PID" + q + " Word docs\\"

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

x = [f for f in os.listdir(folder_path) if f.endswith('.doc')]

pee = 0
# Iterate through each file in the folder
for loc, filename in enumerate(x):
    if pee != loc:
        print("Error around here")
        pee += 1
    if filename.endswith(".doc"):
        # Open the .doc file using win32com
        
        try:
            doc = word.Documents.Open(os.path.join(folder_path, filename))
            for table in doc.Tables:
                table.Delete()
            # Iterate through each paragraph in the document
            for paragraph in doc.Paragraphs:

                # Check if the paragraph is bolded
                if paragraph.Range.Bold and paragraph.Range.Text.strip() != "":


                    current_text = paragraph.Range.Text
                    current_text = r'{}'.format(current_text)
                    
                    paragraph.Range.Text = filename[:-4] + " - " + current_text

                    newtext = paragraph.Range.Text


                
                    # Delete all text before the bolded paragraph


                    # Save the document as a .docx file
                    new_filename = n + "_" + filename[:-4] + " - " + current_text.strip() + ".docx"

                    new_filename.replace("/","")
                    new_filename.replace("\\","")
                    new_filename.replace("<","")
                    new_filename.replace(">","")
                    new_filename.replace(":","")
                    new_filename.replace("|","")
                    new_filename.replace("?","")
                    new_filename.replace("*","")
                    new_filename.replace("\"","")
                    
                    doc.SaveAs(os.path.join(folder_path, new_filename.strip()), FileFormat=16)

                    print(str(loc) + "'th doc: " + new_filename.strip())

                    # Close the document and move on to the next one
                    pee += 1
                    

                    doc.Close()
                    break
        except:
            pee += 1
            doc.Save()
            print("Must manually do " + filename)

        finally:
            continue
        # Quit win32com
x = [f for f in os.listdir(folder_path) if f.endswith('.docx')]

# Iterate through each file in the folder
for filename in x:
    try:
        doc = word.Documents.Open(folder_path + filename)



        for paragraph in doc.Paragraphs:
            if paragraph.Range.Bold and paragraph.Range.Text.strip() != "":
                break
            paragraph.Range.Select()

            word.Selection.Delete()
        for table in doc.Tables:
            table.Delete()
        doc.TrackRevisions = True
        doc.Save()
        doc.Close()
    except:
        print("Fail on " + filename)

for filename in x:
    doc = word.Documents.Open(folder_path + filename)
    try:
        if doc.Paragraphs(1).Range.Text == filename[3:-5]:
            continue
        else:
            print("Check " + filename)
    except:
        print("Should be over now")

word.Quit()
