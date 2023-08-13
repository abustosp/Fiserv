import pdfplumber
import re
import os
import pandas as pd
import openpyxl
import LIB.formatos as fmt
import numpy as np
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo
import time

# Ruta de la carpeta que contiene los archivos PDF
path = askdirectory(title="Seleccionar carpeta con los archivos PDF")

Inicio = time.time()

# Obtener la lista de archivos PDF en la carpeta
files = os.listdir(path)

# Filtrar los archivos que no son PDF
files = [f for f in files if f.endswith(".pdf")]

# Crear la carpeta "Resultados" si no existe
if not os.path.exists(os.path.join(path, "Resultados")):
    os.mkdir(os.path.join(path, "Resultados"))


# Procesar cada archivo PDF
for f in files:
    texto = ""
    # Abrir el archivo PDF
    pdf = pdfplumber.open(os.path.join(path, f))
    # Obtener el texto de todas las páginas
    paginas = pdf.pages
    for pagina in paginas:
        texto += pagina.extract_text()

    # # exportar el texto a un archivo TXT
    # with open(os.path.join("Resultados", f.replace(".pdf", ".txt")), "w", encoding="utf-8") as f_txt:
    #     f_txt.write(texto)

    # Expresión regular para encontrar movimientos
    Re_Movimientos = r"^([-+])\s+(.*?)(?:\s+\$ ([\d,\.]+))?$|el día (\d{2}\/\d{2}\/\d{4}) \$ ([\d,\.]+) Nro\. Liq: (\d+)"
    # Encontrar todos los movimientos
    buscar_movimientos = re.findall(Re_Movimientos, texto, re.MULTILINE)

    # Extraer el CUIT del primer match
    cuit_empresa = re.search(r"CUIT: (\d{2}-\d{8}-\d{1})", texto).group(1)

        # Expresiones regulares para encontrar Total y Pagos
    Total = r"Total presentado: (\d.*)*"
    Pagos = r"Neto de pagos: (\d.*)*"
    # Encontrar el primer match de Total y Pagos
    buscar_total = re.search(Total, texto)
    buscar_pagos = re.search(Pagos, texto)
    # transformar el monto a numérico
    buscar_total = buscar_total.group(1).replace(".", "").replace(",", ".")
    buscar_pagos = buscar_pagos.group(1).replace(".", "").replace(",", ".")

    # transformar los textos de los movimientos a montos numérico de los grupos 3 y 5
    buscar_movimientos = [list(x) for x in buscar_movimientos]  
    for i in range(len(buscar_movimientos)):
        if buscar_movimientos[i][2] != "":
            buscar_movimientos[i][2] = float(buscar_movimientos[i][2].replace(".", "").replace(",", "."))
        if buscar_movimientos[i][4] != "":
            buscar_movimientos[i][4] = float(buscar_movimientos[i][4].replace(".", "").replace(",", "."))
    
    # Crear un DataFrame con los grupos de la expresión regular
    df_movimientos = pd.DataFrame(buscar_movimientos, columns=["Signo", "Concepto", "Monto" , "Fecha", "Importe Neto de Pagos", "Nro. Liq"])

    # Agregar una columna con el CUIT de la empresa
    df_movimientos["CUIT"] = cuit_empresa.replace("-", "")
    df_movimientos["CUIT"] = df_movimientos["CUIT"].astype(np.int64)

    # Multiplicar por -1 los montos negativos
    df_movimientos["Monto"] = df_movimientos.apply(lambda x: x["Monto"] * -1 if x["Signo"] == "-" else x["Monto"], axis=1)

    # Reemplazar los "" por NaN
    df_movimientos = df_movimientos.replace("", float("NaN"))

    # Rellenar las columnas 'Fecha' y 'Importe Neto de Pagos' y 'Nro. Liq' con el valor de la prier fila posterior que no sea NaN
    df_movimientos["Fecha"] = df_movimientos["Fecha"].bfill()
    df_movimientos["Importe Neto de Pagos"] = df_movimientos["Importe Neto de Pagos"].bfill()
    df_movimientos["Nro. Liq"] = df_movimientos["Nro. Liq"].bfill()

    # Eliminar la fila de signo si es NaN
    df_movimientos = df_movimientos.dropna(subset=["Signo"])

    # Crear una Tabla dinámica para agrupar los movimientos por concepto y sumar los montos
    df_TD = df_movimientos.pivot_table(index="Concepto", values="Monto", aggfunc="sum").reset_index()

    # Crear otra tabla dinamica pero que se agrupe por fecha y concepto y se sumen los montos
    df_TD2 = df_movimientos.pivot_table(index=["Fecha", "Concepto"], values="Monto", aggfunc="sum").reset_index()

    # Crear un DataFrame de control donde se sumen los movimientos por su signo (+ o -) y se compare con la diferencia entre Total y Pagos
    df_control = df_movimientos.groupby("Signo").sum().reset_index()

    # Eliminar la columna CUIT 
    df_control = df_control.drop(columns=["CUIT"])

    # Eliminar la columna de 'Fecha' , 'Importe Neto de Pagos' y 'Nro. Liq'
    df_control = df_control.drop(columns=["Fecha", "Importe Neto de Pagos", "Nro. Liq"])

    # Agregar una fila busar_total
    df_control = pd.concat([df_control, pd.DataFrame({"Signo": [""], "Concepto": ["Total Presentado"], "Monto": [buscar_total]})])

    # Agregar una fila buscar_pagos
    df_control = pd.concat([df_control, pd.DataFrame({"Signo": [""], "Concepto": ["Neto de Pagos"], "Monto": [buscar_pagos]})])

    # Crear variable con la suma de los movimientos
    df_control["Monto"] = df_control["Monto"].apply(lambda x: float(x))
    suma_movimientos = df_control["Monto"].sum()

    # Agregar una fila con la suma de los movimientos de pagos
    pagos_Control = df_movimientos[df_movimientos["Signo"] == "-"]["Monto"].sum()
    df_control = pd.concat([df_control, pd.DataFrame({"Signo": [""], "Concepto": ["Pagos"], "Monto": [pagos_Control]})])

    #Calcular la diferencia entre Total y Pagos - la suma de los movimientos que sean negativos
    diferencia = round ( ((float(buscar_total) - float(buscar_pagos)) + df_movimientos[df_movimientos["Signo"] == "-"]["Monto"].sum()) , 2)

    # Agregar una fila de control con la diferencia
    df_control = pd.concat([df_control, pd.DataFrame({"Signo": [""], "Concepto": ["Control"], "Monto": [diferencia]})])

    # Cambiar el concepto de las filas que el signo sea + por "Ingresos"
    df_control.loc[df_control["Signo"] == "+", "Concepto"] = "Ingresos"

    # Cambiar el concepto de las filas que el signo sea - por "Egresos"
    df_control.loc[df_control["Signo"] == "-", "Concepto"] = "Egresos"

    NombreExportación = f.split(".")[0]

    # Exportar a Excel los movimientos y el control
    with pd.ExcelWriter(f"{path}/Resultados/{NombreExportación}.xlsx") as writer:
        df_movimientos.to_excel(writer, sheet_name="Movimientos", index=False)
        df_control.to_excel(writer, sheet_name="Control", index=False)
        df_TD.to_excel(writer, sheet_name="Tabla Dinámica", index=False)
        df_TD2.to_excel(writer, sheet_name="Tabla Dinámica 2", index=False)

    # Aplicar formatos
    workbook = openpyxl.load_workbook(f"{path}/Resultados/{NombreExportación}.xlsx")
    h1 = workbook["Movimientos"]
    h2 = workbook["Control"]
    h3 = workbook["Tabla Dinámica"]
    h4 = workbook["Tabla Dinámica 2"]

    fmt.Aplicar_formato_encabezado(h1)
    fmt.Aplicar_formato_encabezado(h2)
    fmt.Aplicar_formato_encabezado(h3)
    fmt.Aplicar_formato_encabezado(h4)


    fmt.Aplicar_formato_moneda(h1 , 3 , 3)
    fmt.Aplicar_formato_moneda(h1 , 5 , 5)
    
    fmt.Aplicar_formato_moneda(h2 , 3 , 3)

    fmt.Aplicar_formato_moneda(h3 , 2 , 2)

    fmt.Aplicar_formato_moneda(h4 , 3 , 3)
    
    
    fmt.Autoajustar_columnas(h1)
    fmt.Autoajustar_columnas(h2)
    fmt.Autoajustar_columnas(h3)
    fmt.Autoajustar_columnas(h4)

    fmt.Agregar_filtros(h1)
    fmt.Agregar_filtros(h2)
    fmt.Agregar_filtros(h3)
    fmt.Agregar_filtros(h4)

    workbook.save(f"{path}/Resultados/{NombreExportación}.xlsx")

# mostrar mensaje de finalización
Final = time.time()
showinfo("Finalizado", f"Se procesaron con éxito {len(files)} archivos en: {Final - Inicio} Segundos.")
