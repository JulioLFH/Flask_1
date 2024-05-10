"""
from flask import Flask, render_template, redirect, url_for, request, make_response
from markupsafe import escape


app = Flask(__name__)



@app.route('/table1')
def index():
    try:
        # Establecer la conexión con la base de datos HANA
        conn = dbapi.connect(address="192.168.10.14", port=30015, user="SYSTEM", password="@l13xS@pLuR1n2021")

        # Ejecutar una consulta
        cursor = conn.cursor()
        cursor.execute(' SELECT TOP 2 * FROM "SBO_ALIEX_PRODUCCION".OINV ')
        rows = cursor.fetchall()

        # Convertir los resultados a una lista de diccionarios
        data = []
        for row in rows:
            row_dict = {}
            for idx, column in enumerate(cursor.description):
                row_dict[column[0]] = row[idx]
            data.append(row_dict)         

        # Cerrar la conexión
        conn.close()

        return render_template('index.html', data=data)
 
    except Exception as e: 
        return f"Error al conectar a la base de datos: {str(e)}"




#Pages
#@app.route('/')
def index():
    return "CódigoFacilito"

@app.route('/hi')
def hi_world():
    return redirect(url_for('hello'))

#Pages pasar variables
@app.route('/user/<username>')
def user(username):
    return "Welcome user: %s" %escape(username)

@app.route('/post/<int:post_id>')
def show_post(post_id):
    return "Post: %d" % post_id

@app.route('/path/<path:subpath>')
def show_subpath(subpath):
    return "Subpath: %s" % escape(subpath)


@app.route('/hello')
@app.route('/hello/<name>')
def hello(name=None):
    return render_template('hello.html', name=name)

@app.errorhandler(404)
def error(error):
    return render_template('error.html'), 404


@app.route('/')
def index():
    username = request.cookies.get('username')

@app.route('/')
def index():
    resp = make_response(render_template(...))
    resp.set_cookie('username', 'the username')
    return resp



from flask import request

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        if valid_login(request.form['username'],
                       request.form['password']):
            return log_the_user_in(request.form['username'])
        else:
            error='Invalid username/password'
    
    return render_template('login.html',error=error)


if __name__ == "__main__":
    #app.run()
    app.add_url_rule('/',view_func=index)
    app.run(host='0.0.0.0', debug=True)
"""

#librerias
from flask import Flask, render_template, request, redirect, url_for
#from hdbcli import dbapi
import pyodbc
from hdbcli import dbapi


# Configuración de la conexión a SQL Server
server = '192.168.10.50'
database = 'BD_CALIDAD'
username = 'ALIEX'
password = '1234'
conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'


# Establece la conexión a la base de datos SAP HANA
conn2 = dbapi.connect(
    address="192.168.10.14",
    port=30015,
    user="SYSTEM",
    password="@l13xS@pLuR1n2021"
)


app = Flask(__name__)

#Pages
@app.route('/')
@app.route('/index')
def index(name=None):
    return render_template('Mayonesa_Walibi/index.html', name=name)


@app.errorhandler(404)
def error(error):
    return render_template('Mayonesa_Walibi/error.html'), 404

#paginas
@app.route('/Tabla_Analisis')
def Tabla_Analisis():

    conn = pyodbc.connect(conn_str)

    # Aquí podrías incluir el código para obtener los datos de análisis
    # Ejemplo de consulta a la base de datos
    cursor = conn.cursor()
    cursor.execute(""" SELECT * FROM MAYONESA_WALIBI ORDER BY ID_ANALISIS DESC """)
    rows = cursor.fetchall()

    # Convertir los resultados a una lista de diccionarios
    data = []
    for row in rows:
        row_dict = {}
        for idx, column in enumerate(cursor.description):
            row_dict[column[0]] = row[idx]
        data.append(row_dict)
    
    conn.commit()
    cursor.close()   
    return render_template('Mayonesa_Walibi/Tabla_analisis.html', data=data)

@app.route('/Registro_Analisis')
def Registro_Analisis():
    return render_template('Mayonesa_Walibi/Registrar_analisis.html')


#modelado
@app.route('/Detalle')
def Detalle():
    ID = request.args.get('ID')

    conn = pyodbc.connect(conn_str)

    # Aquí podrías incluir el código para obtener los datos de análisis
    # Ejemplo de consulta a la base de datos
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM MAYONESA_WALIBI WHERE ID_ANALISIS = "+ID)
    rows = cursor.fetchall()# Convertir los resultados a una lista de diccionarios
    data = []
    for row in rows:
        row_dict = {}
        for idx, column in enumerate(cursor.description):
            row_dict[column[0]] = row[idx]
        data.append(row_dict)

    #return (data)
    
    conn.commit()
    cursor.close()   
    
    # Aquí puedes utilizar el valor de id para obtener los detalles del elemento con ese ID
    # Por ejemplo, puedes realizar una consulta a la base de datos para obtener los detalles del elemento con ese ID
    return render_template('Mayonesa_Walibi/Detalle_analisis.html', data = data)

@app.route('/Editar')
def Editar():
    ID = request.args.get('ID')

    conn = pyodbc.connect(conn_str)

    # Aquí podrías incluir el código para obtener los datos de análisis
    # Ejemplo de consulta a la base de datos
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM MAYONESA_WALIBI WHERE ID_ANALISIS = "+ID)
    rows = cursor.fetchall()# Convertir los resultados a una lista de diccionarios
    data = []
    for row in rows:
        row_dict = {}
        for idx, column in enumerate(cursor.description):
            row_dict[column[0]] = row[idx]
        data.append(row_dict)

    #return (data)
    
    conn.commit()
    cursor.close()   
    
    # Aquí puedes utilizar el valor de id para obtener los detalles del elemento con ese ID
    # Por ejemplo, puedes realizar una consulta a la base de datos para obtener los detalles del elemento con ese ID
    return render_template('Mayonesa_Walibi/Editar_analisis.html', data = data)

@app.route('/Eliminar')
def Eliminar():

    ID = request.args.get('ID')

    conn = pyodbc.connect(conn_str)

    # Aquí podrías incluir el código para obtener los datos de análisis
    # Ejemplo de consulta a la base de datos
    cursor = conn.cursor()
    cursor.execute("DELETE FROM MAYONESA_WALIBI WHERE ID_ANALISIS = "+ID)
    
    conn.commit()
    cursor.close()
    
    return redirect(url_for('Tabla_Analisis'))



#controlador
@app.route('/guardar_analisis', methods=['POST'])
def guardar_producto():
    if request.method == 'POST':

        conn = pyodbc.connect(conn_str)

        fecha_produccion = request.form['FP']
        fecha_vencimiento = request.form['FV']
        formula = request.form['Formula']
        lote = request.form['Lote']
        ph1 = request.form['PH1']
        ph2 = request.form['PH2']
        ph3 = request.form['PH3']
        ph4 = request.form['PH4']
        ph_prom = request.form['PH-PROM']
        consist1 = request.form['Consist1']
        consist2 = request.form['Consist2']
        color = request.form['Color']
        sabor = request.form['Sabor']
        olor = request.form['Olor']
        apariencia = request.form['Apariencia']
        num_batidas = request.form['batidas']
        num_licuadora = request.form['licuadora']
        presentacion = request.form['presentacion']
        envasadora = request.form['envasadora']
        degustacion = request.form['degustacion']
        observaciones = request.form['observaciones']
        acciones_correctivas = request.form['acciones_correctivas']
        estado = request.form['estado']
        marca = request.form['marca']

        # Insertar los datos en la base de datos SQL Server
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO MAYONESA_WALIBI (
                FECHA_PROD, FECHA_VENCI, FORMULA, LOTE, 
                PH1, PH2, PH3, PH4, PH_PROM, CONSIST_1, CONSIST_2, COLOR, 
                SABOR, OLOR, APARIENCIA, NUM_BATIDAS, NUM_LICUADORA, 
                PRESENTACION, ENVASADORA, DEGUSTACION, OBSERVACIONES, 
                ACCIONES_CORRECTIVAS, ESTADO, MARCA
            ) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            fecha_produccion, fecha_vencimiento, formula, lote, 
            ph1, ph2, ph3, ph4, ph_prom, consist1, consist2, color, 
            sabor, olor, apariencia, num_batidas, num_licuadora, 
            presentacion, envasadora, degustacion, observaciones, 
            acciones_correctivas, estado, marca
        ))
        conn.commit()
        cursor.close()



        return redirect(url_for('Tabla_Analisis'))

@app.route('/actualizar_analisis', methods=['POST'])
def actualizar_analisis():
    if request.method == 'POST':

        conn = pyodbc.connect(conn_str)

        # Obtener el ID_ANALISIS de la solicitud
        ID_ANALISIS = request.form['ID_ANALISIS']

        # Actualizar los demás campos según la solicitud
        fecha_produccion = request.form['FP']
        fecha_vencimiento = request.form['FV']
        formula = request.form['Formula']
        lote = request.form['Lote']
        ph1 = request.form['PH1']
        ph2 = request.form['PH2']
        ph3 = request.form['PH3']
        ph4 = request.form['PH4']
        ph_prom = request.form['PH-PROM']
        consist1 = request.form['Consist1']
        consist2 = request.form['Consist2']
        color = request.form['Color']
        sabor = request.form['Sabor']
        olor = request.form['Olor']
        apariencia = request.form['Apariencia']
        num_batidas = request.form['batidas']
        num_licuadora = request.form['licuadora']
        presentacion = request.form['presentacion']
        envasadora = request.form['envasadora']
        degustacion = request.form['degustacion']
        observaciones = request.form['observaciones']
        acciones_correctivas = request.form['acciones_correctivas']
        estado = request.form['estado']
        marca = request.form['marca']

        # Actualizar los datos en la base de datos SQL Server
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE MAYONESA_WALIBI 
            SET FECHA_PROD=?, FECHA_VENCI=?, FORMULA=?, LOTE=?, 
                PH1=?, PH2=?, PH3=?, PH4=?, PH_PROM=?, CONSIST_1=?, CONSIST_2=?, COLOR=?, 
                SABOR=?, OLOR=?, APARIENCIA=?, NUM_BATIDAS=?, NUM_LICUADORA=?, 
                PRESENTACION=?, ENVASADORA=?, DEGUSTACION=?, OBSERVACIONES=?, 
                ACCIONES_CORRECTIVAS=?, ESTADO=?, MARCA=?
            WHERE ID_ANALISIS=?
        """, (
            fecha_produccion, fecha_vencimiento, formula, lote, 
            ph1, ph2, ph3, ph4, ph_prom, consist1, consist2, color, 
            sabor, olor, apariencia, num_batidas, num_licuadora, 
            presentacion, envasadora, degustacion, observaciones, 
            acciones_correctivas, estado, marca, ID_ANALISIS
        ))

        conn.commit()
        cursor.close()

        return redirect(url_for('Tabla_Analisis'))


#paginas
@app.route('/Consulta_certificado', methods=['GET', 'POST'])
def Consulta_certificado():
    if request.method == 'POST':
        RUC = request.form['RUC']
        LOTE = request.form['LOTE']

        # Ejecutar una consulta
        
        cursor2 = conn2.cursor()
        query_consulta = 'SELECT TOP 1 T2."LicTradNum" AS "RUC", UPPER(T2."CardName") AS "CLIENTE", UPPER(T1."Dscription"),T9."BatchNum" AS "LOTE",T9."Quantity" AS "LOTE_CANTI",T10."ExpDate" as "LOTE_VENCI" FROM "SBO_ALIEX_PRODUCCION"."ODLN" T0 INNER JOIN "SBO_ALIEX_PRODUCCION"."DLN1" T1 ON T1."DocEntry" = T0."DocEntry" INNER JOIN "SBO_ALIEX_PRODUCCION"."OCRD" T2 ON T2."CardCode" = T0."CardCode" left JOIN "SBO_ALIEX_PRODUCCION"."IBT1" T9 ON T9."BaseEntry" = T0."DocEntry" and T9."ItemCode" = T1."ItemCode" and T9."BaseType"=15 left JOIN "SBO_ALIEX_PRODUCCION"."OBTN" T10 ON T10."ItemCode" = T1."ItemCode" and T10."DistNumber" = T9."BatchNum" WHERE T9."BatchNum" = \''+LOTE+'\' AND T2."LicTradNum" = \''+RUC+'\''+ 'AND (UPPER(T1."Dscription") LIKE \'%WALIBI%\' OR UPPER(T1."Dscription") LIKE \'%WALIBÍ%\') AND (UPPER(T1."Dscription") LIKE \'%MAYONESA%\')'
        
        cursor2.execute(query_consulta)
        rows2 = cursor2.fetchall()

        # Obtener la cantidad de filas
        num_filas = cursor2.rowcount

        # Verificar si hay datos
        if num_filas == 0:
            # Si no hay datos, puedes manejarlo aquí
            return render_template('Consulta/Certificado.html', RUC=RUC, LOTE=LOTE, ESTADO = 'NO EXISTE')

        # Convertir los resultados a una lista de diccionarios
        data2 = []
        for row in rows2:
            row_dict = {}
            for idx, column in enumerate(cursor2.description):
                row_dict[column[0]] = row[idx]
            data2.append(row_dict) 

        # Verificar si hay datos
        if not data2:
            # Si no hay datos, puedes manejarlo aquí
            return render_template('Consulta/Certificado.html', RUC=RUC, LOTE=LOTE, ESTADO = 'NO EXISTE')

        CLIENTE_P = data2[0]["CLIENTE"]
        LOTE_P = data2[0]["LOTE"][-6:]

        conn = pyodbc.connect(conn_str)

        # Aquí podrías incluir el código para obtener los datos de análisis
        # Ejemplo de consulta a la base de datos
        cursor = conn.cursor()
        query_consulta2 = "SELECT * FROM MAYONESA_WALIBI WHERE LOTE = '"+LOTE_P+"'"
        cursor.execute(query_consulta2)
        rows = cursor.fetchall()

        # Convertir los resultados a una lista de diccionarios
        data = []
        for row in rows:
            row_dict = {}
            for idx, column in enumerate(cursor.description):
                row_dict[column[0]] = row[idx]
            data.append(row_dict)

        # Verificar si hay datos
        if not data:
            # Si no hay datos, puedes manejarlo aquí
            return render_template('Consulta/Certificado.html', RUC=RUC, LOTE=LOTE, ESTADO = 'NO EXISTE')
        
        PH_PROM_P = data[0]["PH_PROM"]
        FECHA_PROD_P = data[0]["FECHA_PROD"]
        FECHA_VENCI_P = data[0]["FECHA_VENCI"]

        from datetime import date
        # Obtener la fecha actual
        fecha_actual = date.today()
        # Formatear la fecha como una cadena si es necesario
        FECHA_ACTUAL_P = fecha_actual.strftime("%d-%m-%Y")



        import win32com.client
        import os

        # Nombre del archivo Excel
        excel_file = 'templates/Plantilla_certificado/Plantilla_certificado.xlsx'

        # Nombre del archivo PDF
        pdf_file = 'C:/Users/Administrador.SRV-ALIEX-D/Desktop/FLASK/static/certificados/Certificado_' + RUC + '_' + LOTE + '.pdf'

        # Validar si el archivo existe
        if os.path.exists(pdf_file):
            # Eliminar el archivo
            os.remove(pdf_file)
            print("El archivo PDF ha sido eliminado exitosamente.")
        else:
            print("El archivo PDF no existe.")

        # Crear una instancia de Excel
        excel = win32com.client.Dispatch("Excel.Application")

        # Abrir el archivo Excel
        workbook = excel.Workbooks.Open(os.path.abspath(excel_file))

        # Seleccionar la hoja de trabajo activa
        sheet = workbook.ActiveSheet

        # Modificar el valor de las celdas
        sheet.Range('C14').Value = str(CLIENTE_P)
        sheet.Range('F14').Value = "Fecha de emisión certificado:" + str(FECHA_ACTUAL_P)
        sheet.Range('C17').Value = str(LOTE_P)
        sheet.Range('C18').Value = str(FECHA_PROD_P)
        sheet.Range('C19').Value = str(FECHA_VENCI_P)
        sheet.Range('C23').Value = str(PH_PROM_P)

        # Convertir a PDF
        workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_file)

        # Cerrar Excel
        workbook.Close(SaveChanges=True)
        excel.Quit()

        # Cerrar el cursor y la conexión
        cursor2.close()  



        
        return render_template('Consulta/Certificado.html', RUC=RUC, LOTE=LOTE, ESTADO = 'CORRECTO')
    
    else:
        RUC = ''
        LOTE = ''
        # Aquí puedes procesar los datos (por ejemplo, generar el certificado) antes de redirigir los datos en la URL
        return render_template('Consulta/Certificado.html', RUC=RUC, LOTE=LOTE)
        # Aquí simplemente redirigimos a certificado.html pasando



if __name__ == '__main__':
    app.run(host='0.0.0.0',port=9090, debug=True)
