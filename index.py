import os
import comtypes.client

def docx_to_pdf(docx_directory, pdf_directory):
    # Crear una instancia en word
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False

    # Crear el directorio PDF si no existe
    if not os.path.exists(pdf_directory):
        os.makedirs(pdf_directory)

    # Convertir las rutas en rutas absolutas
    docx_directory = os.path.abspath(docx_directory)
    pdf_directory = os.path.abspath(pdf_directory)

    # Verificar si el directorio DOCX existe
    if not os.path.exists(docx_directory):
        raise FileNotFoundError(f"El directorio {docx_directory} no existe.")

    # Procesar archivos DOCX
    for docx_file in os.listdir(docx_directory):
        if docx_file.endswith(".docx"):
            docx_path = os.path.join(docx_directory, docx_file)
            pdf_path = os.path.join(pdf_directory, docx_file.replace(".docx", ".pdf"))

            if os.path.exists(pdf_path):
                print(f"Actualizando el archivo: {pdf_path}")
            else:
                print(f"Creando el archivo: {pdf_path}")

            # print("Intentando abrir:", docx_path)
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
    
    print("Proceso finalizado.")
    word.Quit()

def main():

    # Directorios de entrada y salida
    input_directory = r'./WORDs'
    output_directory = r'./PDFs'

    # Llamar a la función para convertir los archivos
    docx_to_pdf(input_directory, output_directory)

if __name__ == "__main__":
    main()
