import os
import docx
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, Entry, Label, Button
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import shutil

def extrage_text_din_docx(docx_path):
    text = ""
    doc = docx.Document(docx_path)
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

def extrage_text_din_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as pdf_document:
        for pagina_numar in range(pdf_document.page_count):
            pagina = pdf_document[pagina_numar]
            text += pagina.get_text()

    return text

def calculeaza_similaritate(document1, document2):
    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform([document1, document2])
    similarity_matrix = cosine_similarity(tfidf_matrix, tfidf_matrix)
    return similarity_matrix[0, 1]

def organizeaza_documente(director_input, director_output, prag_similaritate=0.7):
    documente = {}

    for nume_fisier in os.listdir(director_input):
        cale_absoluta = os.path.join(director_input, nume_fisier)
        extensie = os.path.splitext(nume_fisier)[1].lower()

        if extensie == '.docx':
            text_document = extrage_text_din_docx(cale_absoluta)
        elif extensie == '.pdf':
            text_document = extrage_text_din_pdf(cale_absoluta)
        elif extensie == '.txt':
            with open(cale_absoluta, 'r', encoding='utf-8') as file:
                text_document = file.read()
        else:
            continue  # Ignora fisierele cu extensii neacceptate

        documente[nume_fisier] = text_document

    if not documente:
        return  # Nu există documente de procesat

    grupuri = {}
    for nume_fisier1, document1 in documente.items():
        grup_gasit = False
        for grup, fisiere in grupuri.items():
            for nume_fisier2 in fisiere:
                document2 = documente[nume_fisier2]
                similaritate = calculeaza_similaritate(document1, document2)
                if similaritate > prag_similaritate:
                    grupuri[grup].append(nume_fisier1)
                    grup_gasit = True
                    break

            if grup_gasit:
                break

        if not grup_gasit:
            grupuri[nume_fisier1] = [nume_fisier1]

    for index, (grup, fisiere) in enumerate(grupuri.items(), start=1):
        director_destinatie = os.path.join(director_output, f"Grup{index}")
        if not os.path.exists(director_destinatie):
            os.makedirs(director_destinatie)

        for nume_fisier in fisiere:
            cale_destinatie = os.path.join(director_destinatie, nume_fisier)
            shutil.copy(os.path.join(director_input, nume_fisier), cale_destinatie)

    return grupuri

def selecteaza_director_input():
    director_input = filedialog.askdirectory()
    entry_director_input.delete(0, tk.END)
    entry_director_input.insert(0, director_input)

def selecteaza_director_output():
    director_output = filedialog.askdirectory()
    entry_director_output.delete(0, tk.END)
    entry_director_output.insert(0, director_output)

def organizeaza_cu_interfata():
    lbl_status = tk.Label(root, text="")
    lbl_status.grid(row=3, column=1)
    director_input = entry_director_input.get()
    director_output = entry_director_output.get()
    organizeaza_documente(director_input, director_output)
    lbl_status.config(text="Organizare finalizată!")

# Creare interfață grafică
root = tk.Tk()
root.title("Organizare Documente")
root.config(background='#7abfeb')

# Etichete și intrări pentru selectarea directorului de intrare și directorului de ieșire
lbl_director_input = tk.Label(root, text="Director Intrare:")
lbl_director_input.grid(row=0, column=0, pady=10, padx=10)
entry_director_input = tk.Entry(root, width=50)
entry_director_input.grid(row=0, column=1, pady=10, padx=10)
btn_selecteaza_input = tk.Button(root, text="Selectează" , bg='red', activebackground='green' , command=selecteaza_director_input)
btn_selecteaza_input.grid(row=0, column=2, pady=10, padx=10)

lbl_director_output = tk.Label(root, text="Director Ieșire:")
lbl_director_output.grid(row=1, column=0, pady=10, padx=10)
entry_director_output = tk.Entry(root, width=50)
entry_director_output.grid(row=1, column=1, pady=10, padx=10)
btn_selecteaza_output = tk.Button(root, text="Selectează",bg='red', activebackground='green', command=selecteaza_director_output)
btn_selecteaza_output.grid(row=1, column=2, pady=10, padx=10)

# Buton pentru organizarea documentelor
btn_organizeaza = tk.Button(root, text="Organizează Documente", command=organizeaza_cu_interfata, width=20)
btn_organizeaza.grid(row=2, column=1, pady=20)

# Etichetă pentru afișarea stării
lbl_status = tk.Label(root, text="")
lbl_status.grid(row=3, column=1)

root.mainloop()
