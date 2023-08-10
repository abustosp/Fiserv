import pdfplumber
import re
import os
import pandas as pd

path = "Muestras"
files = os.listdir(path)

# Filtrar las que no son .pdf
files = [f for f in files if f.endswith(".pdf")]

# Crear la carpeta Resultados si no existe
if not os.path.exists("Resultados"):
    os.mkdir("Resultados")

for f in files:

    # Abrir el archivo PDF
    pdf = pdfplumber.open(os.path.join(path, f))

    # Obtener el texto de todas las páginas
    paginas = pdf.pages
    texto = ""
    for pagina in paginas:
        texto += pagina.extract_text()

    Re_Movimientos = "^([-+])\s+(.*?)(?:\s+\$\s*([\d,\.]+))$"

    Total = "Total presentado: (\d.*)*"
    Pagos = "Neto de pagos: (\d.*)*"

    # Encontrar todos los movimientos
    buscar_movimientos = re.findall(Re_Movimientos, texto, re.MULTILINE)

    # Encontrar el primer match de Total y Pagos
    buscar_total = re.search(Total, texto)
    buscar_pagos = re.search(Pagos, texto)

    # Crear un DataFrame con los movimientos, Total y Pagos
    df = pd.DataFrame(buscar_movimientos, columns=["Signo", "Concepto", "Monto"])

    # Convertir el monto a numérico (donde el separador de miles es el punto y el decimal la coma)
    df["Monto"] = df["Monto"].apply(lambda x: float(x.replace(".", "").replace(",", ".")))

    # Multiplicar por -1 los montos negativos
    df["Monto"] = df.apply(lambda x: x["Monto"] * -1 if x["Signo"] == "-" else x["Monto"], axis=1)

    # Crear un df de control donde se sumen los movimientos por su signo (+ o -) y se compare con la diferencia entre Total y Pagos
    df_control = df.groupby("Signo").sum().reset_index()

    # Agregar una fila con el Total y Pagos
    df_control = df_control.append({"Concepto": "Total", "Monto": buscar_total.group(1)}, ignore_index=True)
    df_control = df_control.append({"Concepto": "Pagos", "Monto": buscar_pagos.group(1)}, ignore_index=True)

    # Agregar una fila de control con la diferencia entre Total y Pagos - la suma de los movimientos
    df_control = df_control.append({"Concepto": "Diferencia", "Monto": float(buscar_total.group(1)) - float(buscar_pagos.group(1)) - df_control["Monto"].sum()}, ignore_index=True)

    # Exportar a Excel los movimientos y el control
    with pd.ExcelWriter(f"Resultados/{f}.xlsx") as writer:
        df.to_excel(writer, sheet_name="Movimientos", index=False)
        df_control.to_excel(writer, sheet_name="Control", index=False)    


        
    # print(texto)

    print("Total: ", buscar_total.group(1))
    print("Pagos: ", buscar_pagos.group(1))
    print("""------------------------------------------------
        Movimientos:
        """)

    for i in buscar_movimientos:
        print(f"""{i[0]} _ {i[1]} _ {i[2]}""")

"""
Explicación del RegEx
1. `^([-+])` : Esto captura el primer signo (+0 -) al principio de una línea y lo coloca en un
grupo de captura.
2. `\s+` : Coincide con uno o más espacios en blanco después del signo.
3. `(.*?)` : Captura el concepto, que puede contener cualquier cantidad de caracteres, incluso ninguno. El uso de `?` hace que la captura sea no codiciosa.
(2: t Xd , : Esto captura el monto en dólares ($), que puede contener
4.`(?:\s+\$\s*([\d,\.]+))$` : Esto captura el monto en dólares ($), que puede contener dígitos, comas y puntos. El uso de `(?: ... )` crea un grupo no capturador para esta parte. `\s+` coincide con uno o más espacios en blanco, `\$\s*` coincide con el símbolo $ y cualquier cantidad de espacios en blanco después, `([\d,\.]+)` captura los dígitos, comas y puntos en el monto, y `$` coincide con el final de la línea.

"""