import win32com.client
import docx2txt

# convert .doc to .docx
word = win32com.client.Dispatch("Word.application")
wordDoc = word.Documents.Open('daftar5.doc')
wordDoc.SaveAs2('daftar5.docx', FileFormat=16)
wordDoc.Close()

# convert .docx to .txt
text = docx2txt.process('daftar5.docx')
file = open('daftar5.txt', 'w', encoding='UTF-8')
file.write(text)
file.close()
