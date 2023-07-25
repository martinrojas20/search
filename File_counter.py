import os
import collections
import tkinter as tk
from tkinter import filedialog, messagebox
import PyPDF2
from reportlab.pdfgen import canvas




def obtener_numero_paginas_pdf(archivo_pdf):
    with open(archivo_pdf, 'rb') as archivo:
        lector_pdf = PyPDF2.PdfReader(archivo)
        numero_paginas = len(lector_pdf.pages)
        return numero_paginas


def contar_dir_y_pag(ruta):    
    total_directorios_81 = 0
    total_directorios_54 = 0   
    total_paginas_pdf = 0
    for directorio_actual, directorios, archivos in os.walk(ruta):
        print(f'actual: {directorio_actual} dir: {directorios}')
        for directorio in directorios:
            if directorio.startswith('81'):
                total_directorios_81 += 1
                directorio_completo = os.path.join(directorio_actual, directorio)                
                for subdirectorio, _, archivos_subdir in os.walk(directorio_completo):
                    for archivo in archivos_subdir:
                            if archivo.endswith('.pdf'):
                                archivo_completo = os.path.join(subdirectorio, archivo)  
                                cantidad_pag = obtener_numero_paginas_pdf( archivo_completo)            
                                total_paginas_pdf = total_paginas_pdf + cantidad_pag  

            elif directorio.startswith('54'):
                total_directorios_54 +=  1
                directorio_completo = os.path.join(directorio_actual, directorio)
                print(f'DC: {directorio_completo}')
                for subdirectorio, _, archivos_subdir in os.walk(directorio_completo):
                    print(f'archivos: {archivos_subdir}')
                    for archivo in archivos_subdir:
                            if archivo.endswith('.pdf'):
                                archivo_completo = os.path.join(subdirectorio, archivo)  
                                cantidad_pag = obtener_numero_paginas_pdf( archivo_completo)            
                                total_paginas_pdf = total_paginas_pdf + cantidad_pag   
            
    
    return total_directorios_81, total_directorios_54, total_paginas_pdf



def seleccionar_ruta():
    ruta = filedialog.askdirectory()
    if ruta:
        resultado = contar_dir_y_pag(ruta)        
        guardar_pdf(resultado) 

    else:
        messagebox.showwarning("Advertencia", "No se seleccionó ninguna ruta.")



def guardar_pdf(resultado):
    ruta_guardar = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Archivo PDF", "*.pdf")])
    if ruta_guardar:
        nombre_archivo = os.path.basename(ruta_guardar)
        
        pdf = canvas.Canvas(ruta_guardar)
        pdf.setFont("Helvetica", 12)
        
        pdf.drawString(150, 700, f"REPORTE DE DIRECTORIOS Y TOTAL PAGINAS PDFS: ")
        pdf.drawString(50, 650, f"Total de directorios que empiezan con '81': {resultado[0]}")
        pdf.drawString(50, 630, f"Total de directorios que empiezan con '54': {resultado[1]}")
        pdf.drawString(50, 610, f"Total de páginas en archivos PDF: {resultado[2]}")

        pdf.save()
        print(f"Archivo PDF generado y guardado en: {ruta_guardar}")


# Crear ventana principal
ventana = tk.Tk()
ventana.title("Contador de Directorios y Páginas en PDF")
ventana.geometry("400x200")

# Etiqueta y botón para seleccionar ruta
etiqueta = tk.Label(ventana, text="Seleccione una ruta de directorio:")
etiqueta.pack(pady=10)

boton_ruta = tk.Button(ventana, text="Seleccionar ruta", command=seleccionar_ruta)
boton_ruta.pack(pady=5)

# Ejecutar ventana
ventana.mainloop()