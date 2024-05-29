import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import json
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import fitz
import base64
from spire.pdf.common import *

from spire.pdf import *
from spire.pdf import PdfDocument



def convertPdfToXml(sourceFile, destinationFile):
    try:
        with pdfplumber.open(sourceFile) as pdf:
            root = ET.Element("root")
            for i, page in enumerate(pdf.pages):
                page_element = ET.SubElement(root, f"page_{i + 1}")

                # Metni çıkar ve XML'e ekle
                page_text = page.extract_text()
                if page_text:
                    text_element = ET.SubElement(page_element, "text")
                    text_element.text = page_text

                # Tabloları çıkar ve XML'e ekle
                tables = page.extract_tables()
                if tables:
                    tables_element = ET.SubElement(page_element, "tables")
                    for k, table in enumerate(tables):
                        table_element = ET.SubElement(tables_element, f"table_{k + 1}")
                        for row in table:
                            row_element = ET.SubElement(table_element, "row")
                            for cell in row:
                                cell_element = ET.SubElement(row_element, "cell")
                                cell_element.text = str(cell)

                # Resimleri çıkar ve XML'e ekle
                images = page.images
                if images:
                    images_element = ET.SubElement(page_element, "images")
                    for img in images:
                        img_element = ET.SubElement(images_element, "image")
                        img_element.text = str(img)

        tree = ET.ElementTree(root)
        tree.write(destinationFile)
        messagebox.showinfo("Başarılı", f"Sonuç dosyası \"{destinationFile}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))


def convertPdfToJson(sourceFile, destinationFile):
    try:
        pdf_document = fitz.open(sourceFile)
        data = {"pages": []}

        with pdfplumber.open(sourceFile) as pdf:
            for i, page in enumerate(pdf.pages):
                page_data = {
                    "page_number": i + 1,
                    "text": page.extract_text(),
                    "tables": page.extract_tables(),
                    "images": []
                }

                page_objects = pdf_document.load_page(i)
                images = page_objects.get_images(full=True)
                for img_info in images:
                    xref = img_info[0]
                    base_image = pdf_document.extract_image(xref)
                    image_data = {
                        "image_number": len(page_data["images"]) + 1,
                        "image_extension": base_image["ext"]
                    }
                    page_data["images"].append(image_data)

                data["pages"].append(page_data)

        with open(destinationFile, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)

        messagebox.showinfo("Başarılı", f"Sonuç dosyası \"{destinationFile}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))


def convertPdfToHtml(sourceFile, destinationFile):
    try:
        # Create a PdfDocument object
        doc = PdfDocument()

        # Load the PDF document
        doc.LoadFromFile(sourceFile)

        # Disable embedding SVG
        doc.ConvertOptions.SetPdfToHtmlOptions(True, False, 1, False)
        doc.SaveToFile(destinationFile, FileFormat.HTML)

        # Close the document
        doc.Close()

        messagebox.showinfo("Başarılı", f"Sonuç dosyası \"{destinationFile}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))










def convertPdfToText(sourceFile, destinationFile):
    try:
        with pdfplumber.open(sourceFile) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text

                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        text += "\n\nTablo:\n"
                        for row in table:
                            text += "\t".join([str(cell) for cell in row]) + "\n"

                images = page.images
                if images:
                    for img in images:
                        text += f"\n\nResim: {img}\n"

        with open(destinationFile, "w", encoding="utf-8") as file:
            file.write(text)

        messagebox.showinfo("Başarılı", f"Sonuç dosyası \"{destinationFile}\" dosyası olarak kaydedildi.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))


def browsePdfFile():
    file_path = filedialog.askopenfilename(filetypes=[("PDF dosyaları", "*.pdf")])
    if file_path:
        entry_pdf_path.delete(0, tk.END)
        entry_pdf_path.insert(0, file_path)


def saveFile(conversion_function, entry_output_path, output_extension):
    file_path = filedialog.asksaveasfilename(defaultextension=output_extension, filetypes=[
        (f"{output_extension.upper()} dosyaları", f"*{output_extension}")])
    if file_path:
        entry_output_path.delete(0, tk.END)
        entry_output_path.insert(0, file_path)
        conversion_function(entry_pdf_path.get(), file_path)


def saveXmlFile():
    saveFile(convertPdfToXml, entry_xml_path, ".xml")


def saveJsonFile():
    saveFile(convertPdfToJson, entry_json_path, ".json")


def saveHtmlFile():
    saveFile(convertPdfToHtml, entry_html_path, ".html")


def saveTextFile():
    saveFile(convertPdfToText, entry_text_path, ".txt")


# Arayüz bileşenlerinin oluşturulması
root = tk.Tk()
root.title("PDF Dönüştürücü")

label_pdf_path = tk.Label(root, text="PDF Dosya Yolu:")
label_pdf_path.grid(row=0, column=0, padx=10, pady=10, sticky="e")

entry_pdf_path = tk.Entry(root, width=50)
entry_pdf_path.grid(row=0, column=1, padx=10, pady=10)

button_browse_pdf = tk.Button(root, text="Göz at...", command=browsePdfFile)
button_browse_pdf.grid(row=0, column=2, padx=10, pady=10)

label_xml_path = tk.Label(root, text="XML Kaydet:")
label_xml_path.grid(row=1, column=0, padx=10, pady=10, sticky="e")

entry_xml_path = tk.Entry(root, width=50)
entry_xml_path.grid(row=1, column=1, padx=10, pady=10)

button_save_xml = tk.Button(root, text="Kaydet...", command=saveXmlFile)
button_save_xml.grid(row=1, column=2, padx=10, pady=10)

label_json_path = tk.Label(root, text="JSON Kaydet:")
label_json_path.grid(row=2, column=0, padx=10, pady=10, sticky="e")

entry_json_path = tk.Entry(root, width=50)
entry_json_path.grid(row=2, column=1, padx=10, pady=10)

button_save_json = tk.Button(root, text="Kaydet...", command=saveJsonFile)
button_save_json.grid(row=2, column=2, padx=10, pady=10)

label_html_path = tk.Label(root, text="HTML Kaydet:")
label_html_path.grid(row=3, column=0, padx=10, pady=10, sticky="e")

entry_html_path = tk.Entry(root, width=50)
entry_html_path.grid(row=3, column=1, padx=10, pady=10)

button_save_html = tk.Button(root, text="Kaydet...", command=saveHtmlFile)
button_save_html.grid(row=3, column=2, padx=10, pady=10)

label_text_path = tk.Label(root, text="Metin Kaydet:")
label_text_path.grid(row=4, column=0, padx=10, pady=10, sticky="e")

entry_text_path = tk.Entry(root, width=50)
entry_text_path.grid(row=4, column=1, padx=10, pady=10)

button_save_text = tk.Button(root, text="Kaydet...", command=saveTextFile)
button_save_text.grid(row=4, column=2, padx=10, pady=10)

root.mainloop()