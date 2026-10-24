# Librerías
import pandas as pd
from tabulate import tabulate
from fpdf import FPDF

# Datos generales
nombre_estudiante = "Santiago Arango"
archivo_excel = "datos_ventas_santiago.xlsx"
archivo_pdf = "reporte_ejecutivo_ventas_santiago.pdf"

# Productos y sus precios
productos_disponibles = {
    "Tablet 10'": 350, 
    "Webcam HD": 45,
    "Disco Duro Externo": 90,
    "Monitor Curvo 27'": 280,
    "Parlantes Bluetooth": 70,
    "Router WiFi 6": 110,
    "Silla Ergonómica": 220
}

# Ventas fijas: (vendedor, producto, cantidad)
ventas = [
    ("Mateo", "Router WiFi 6", 4),
    ("Laura", "Webcam HD", 6),
    ("Laura", "Router WiFi 6", 1),
    ("Camila", "Disco Duro Externo", 3),
    ("Camila", "Monitor Curvo 27'", 2),
    ("Sebastián", "Monitor Curvo 27'", 1),
    ("Valentina", "Tablet 10'", 4),
    ("Valentina", "Disco Duro Externo", 3),
    ("Camila", "Monitor Curvo 27'", 5),
    ("Mateo", "Router WiFi 6", 4),
    ("Sebastián", "Tablet 10'", 2),
    ("Mateo", "Router WiFi 6", 1),
    ("Sebastián", "Parlantes Bluetooth", 2),
    ("Camila", "Webcam HD", 5),
    ("Sebastián", "Disco Duro Externo", 5),
    ("Camila", "Monitor Curvo 27'", 4),
    ("Laura", "Silla Ergonómica", 1),
    ("Sebastián", "Tablet 10'", 1),
    ("Laura", "Silla Ergonómica", 2),
    ("Valentina", "Router WiFi 6", 4),
    ("Camila", "Disco Duro Externo", 6),
    ("Sebastián", "Tablet 10'", 4),
    ("Sebastián", "Silla Ergonómica", 5),
    ("Sebastián", "Silla Ergonómica", 5),
    ("Sebastián", "Parlantes Bluetooth", 6),
]

# Armar cada venta con su precio y total
datos = []
for i, (vendedor, producto, cantidad) in enumerate(ventas, start=1):
    precio = productos_disponibles[producto]
    total = precio * cantidad
    datos.append([i, vendedor, producto, cantidad, precio, total])

# Crear la tabla con pandas
columnas = ["ID", "Vendedor", "Producto", "Cantidad", "Precio Unit.", "Total"]
df = pd.DataFrame(datos, columns=columnas)

# Calcular resultados
total_general = df["Total"].sum()
promedio_venta = df["Total"].mean()
ventas_vendedor = df.groupby("Vendedor")["Total"].sum()
vendedor_top = ventas_vendedor.idxmax()
venta_top_valor = ventas_vendedor.max()


# Opcion 1: ver todas las ventas
def ver_ventas():
    print(tabulate(df, headers='keys', tablefmt='grid', showindex=False))


# Opcion 2: resumen general
def ver_resumen():
    print(f"Total General de Ventas: ${total_general}")
    print(f"Promedio por Venta: ${promedio_venta:.2f}")
    print(f"Mejor Vendedor: {vendedor_top} con ${venta_top_valor}")


# Opcion 3: total vendido por cada vendedor
def ver_por_vendedor():
    tabla = ventas_vendedor.sort_values(ascending=False).reset_index()
    print(tabulate(tabla, headers=["Vendedor", "Total"], tablefmt='grid', showindex=False))


# Opcion 4: guardar las ventas en Excel
def generar_excel():
    try:
        df.to_excel(archivo_excel, index=False)
        print(f"Archivo {archivo_excel} generado con éxito.")
    except PermissionError:
        print("Cierre el Excel e intente de nuevo")


# Opcion 5: crear el reporte en PDF
def generar_pdf():
    pdf = FPDF()
    pdf.add_page()

    # Título y nombre
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(200, 10, text="REPORTE EJECUTIVO DE VENTAS", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.set_font("Helvetica", size=12)
    pdf.cell(200, 10, text=f"Estudiante: {nombre_estudiante}", new_x="LMARGIN", new_y="NEXT", align='L')
    pdf.ln(10)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(200, 10, text="Detalle de Operaciones:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=10)

    # Encabezados de la tabla
    pdf.cell(35, 10, "Vendedor", 1)
    pdf.cell(60, 10, "Producto", 1)
    pdf.cell(30, 10, "Cantidad", 1)
    pdf.cell(40, 10, "Total", 1, new_x="LMARGIN", new_y="NEXT")

    # Filas de la tabla
    for index, row in df.iterrows():
        pdf.cell(35, 10, str(row["Vendedor"]), 1)
        pdf.cell(60, 10, str(row["Producto"]), 1)
        pdf.cell(30, 10, str(row["Cantidad"]), 1)
        pdf.cell(40, 10, f"${row['Total']}", 1, new_x="LMARGIN", new_y="NEXT")

    # Resumen final
    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(200, 8, text=f"Total General: ${total_general}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(200, 8, text=f"Mejor Vendedor: {vendedor_top} (${venta_top_valor})", new_x="LMARGIN", new_y="NEXT")

    try:
        pdf.output(archivo_pdf)
        print(f"Archivo {archivo_pdf} generado con éxito.")
    except PermissionError:
        print("Cierre el PDF e intente de nuevo")


# Menu interactivo (se repite hasta elegir 6)
def menu():
    opcion = ""
    while opcion != "6":
        print()
        print("====== MENU DE VENTAS ======")
        print("1. Ver todas las ventas")
        print("2. Ver resumen")
        print("3. Ventas por vendedor")
        print("4. Generar Excel")
        print("5. Generar PDF")
        print("6. Salir")
        opcion = input("Elija una opcion: ").strip()
        print()

        if opcion == "1":
            ver_ventas()
        elif opcion == "2":
            ver_resumen()
        elif opcion == "3":
            ver_por_vendedor()
        elif opcion == "4":
            generar_excel()
        elif opcion == "5":
            generar_pdf()
        elif opcion == "6":
            print("Hasta luego")
        else:
            print("Opcion no valida, escriba un numero del 1 al 6")


menu()   # arranca el programa
