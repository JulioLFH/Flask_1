import win32com.client
import openpyxl

# Nombre del archivo Excel
excel_file = 'Plantilla_certificado.xlsx'

# Abrir el archivo Excel
workbook = openpyxl.load_workbook(excel_file)

# Seleccionar la hoja de trabajo
sheet = workbook.active

# Modificar el valor de la celda C17
sheet['C17'] = '122-25'

# Guardar los cambios en el archivo Excel
workbook.save(excel_file)


def excel_to_pdf(excel_file, pdf_file):
    # Crear una instancia de Excel
    excel = win32com.client.Dispatch("Excel.Application")

    # Abrir el archivo Excel
    workbook = excel.Workbooks.Open(excel_file)

    # Convertir a PDF
    workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_file)

    # Cerrar Excel
    workbook.Close(SaveChanges=False)
    excel.Quit()

# Ejemplo de uso
excel_file = 'Plantilla_certificado.xlsx'
pdf_file = 'Certificado.pdf'
excel_to_pdf(excel_file, pdf_file)