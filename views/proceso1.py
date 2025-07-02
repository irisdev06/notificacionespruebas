from openpyxl.styles import Border, Side, PatternFill
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
import matplotlib.pyplot as plt
from io import BytesIO
import streamlit as st
import pandas as pd
import calendar
import locale


def subir_archivo():
    archivo = st.file_uploader("Sube un archivo (.xlsx o .csv)", type=["xlsx", "csv"])
    
    if archivo is not None:
        try:
            nombre_archivo = archivo.name.lower()

            if nombre_archivo.endswith(".xlsx"):
                xls = pd.ExcelFile(archivo)
                hojas = xls.sheet_names

                if "DTO" in hojas and "PCL" in hojas:
                    st.success("¡Archivo Excel válido! Se encontraron las hojas DTO y PCL.")
                    return archivo, "xlsx"
                else:
                    if "DTO" not in hojas:
                        st.error("La hoja 'DTO' no se encuentra en el archivo.")
                    if "PCL" not in hojas:
                        st.error("La hoja 'PCL' no se encuentra en el archivo.")
                    return None, None

            elif nombre_archivo.endswith(".csv"):
                df = pd.read_csv(archivo)
                if "DTO" in df.columns and "PCL" in df.columns:
                    st.success("¡Archivo CSV válido! Se encontraron las columnas DTO y PCL.")
                    return archivo, "csv"
                else:
                    st.warning("El archivo CSV no contiene columnas llamadas 'DTO' y 'PCL'.")
                    return None, None

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            return None, None

    return None, None

colores = ['#FFB897', '#B8E6A7', '#809bce', "#64a09d", '#CBE6FF', '#E6E6FA']

# ------------------------------------------------------------------------------ GRÁFICOS BARRAS---------------------------------------------------------------------------------
def graficas_barras(df, colores, nombre_hoja):
    # Agrupar por MES y NOTIFICADOR y contar los registros
    conteo = df.groupby(['MES', 'NOTIFICADOR']).size().unstack(fill_value=0)
    conteo.index = conteo.index.map(lambda m: calendar.month_name[m].capitalize())

    # Calcular el tamaño dinámico del gráfico en función de los datos
    num_meses = len(conteo)
    num_notificadores = len(conteo.columns)

    # Establecer tamaño de la figura dinámicamente
    fig_width = max(12, num_meses * 1.2)  # Aumentamos el tamaño del gráfico
    fig_height = max(8, num_notificadores * 1.0)

    # Crear gráfico de barras agrupadas
    ax = conteo.plot(kind='bar', figsize=(fig_width, fig_height), color=colores)

    # Personalizar la gráfica
    ax.set_xlabel('Mes')
    ax.set_ylabel('Número de Datos')

    # Ajustar la leyenda para que siempre esté visible
    ax.legend(title='NOTIFICADOR', bbox_to_anchor=(1.2, 1), loc='upper left', fontsize=10)

    # Añadir los números encima de cada barra
    for p in ax.patches:
        ax.annotate(f'{p.get_height()}', 
                    (p.get_x() + p.get_width() / 2., p.get_height()), 
                    xytext=(0, 5),  # Desplazamiento del texto
                    textcoords='offset points',
                    ha='center', va='bottom', fontsize=10, color='black')

    # Guardar la gráfica como imagen
    grafico_path = f"{nombre_hoja}_grafico_barras.png"
    ax.get_figure().savefig(grafico_path, transparent=True, bbox_inches="tight")

    # Devolver la ruta de la imagen guardada
    return grafico_path




# ------------------------------------------------------------------------------ GRÁFICOS PASTEL ---------------------------------------------------------------------------------
def graficas_pastel(df, nombre_hoja):
    # Agrupar por MES y contar los registros (sin usar NOTIFICADOR)
    conteo = df.groupby('MES').size()

    # Convertir los meses a nombres completos
    conteo.index = conteo.index.map(lambda m: calendar.month_name[m].capitalize())  

    # Crear gráfico de pastel
    fig, ax = plt.subplots(figsize=(8, 8))
    ax.pie(conteo, labels=conteo.index, autopct='%1.1f%%', startangle=90, colors=colores)

    # Personalizar la gráfica
    ax.legend(title='Meses', loc='center left', bbox_to_anchor=(1.05, 0.5), fontsize=10)

    # Guardar la gráfica como imagen
    grafico_path = f"{nombre_hoja}_grafico_pastel.png"
    plt.tight_layout()
    plt.savefig(grafico_path, transparent=True, bbox_inches="tight")

    # Devolver la ruta de la imagen guardada
    return grafico_path
# ------------------------------------------------------------------------------ TABLAS ---------------------------------------------------------------------------------

def generar_tablas_dto_y_pcl(archivo_subido):
    # Configurar idioma a español para nombres de meses
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'es_CO.UTF-8')
        except:
            locale.setlocale(locale.LC_TIME, '')

    archivo_subido.seek(0)
    archivo_bytes = BytesIO(archivo_subido.read())
    libro = load_workbook(archivo_bytes)

    def crear_hoja(nombre_hoja, df):
        # Procesar los datos para la tabla
        df['MES'] = df['FECHA_VISADO'].dt.month
        conteo = df.groupby('MES').size().reset_index(name='TOTAL')
        conteo['MES'] = conteo['MES'].apply(lambda m: calendar.month_name[m].capitalize())
        total_general = conteo['TOTAL'].sum()
        conteo['PORCENTAJE'] = (conteo['TOTAL'] / total_general * 100).round(2).astype(str) + '%'

        fila_total = pd.DataFrame({
            'MES': ['Total general'],
            'TOTAL': [total_general],
            'PORCENTAJE': ['100.0%']
        })
        tabla_final = pd.concat([conteo, fila_total], ignore_index=True)

        # Crear o reemplazar la hoja de Excel
        if nombre_hoja in libro.sheetnames:
            del libro[nombre_hoja]
        hoja = libro.create_sheet(nombre_hoja)

        borde = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000")
        )
        fondo_gris = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        hoja['A1'] = "FECHA VISADO"
        hoja['B1'] = "TOTAL"
        hoja['C1'] = "PORCENTAJE"

        for celda in ['A1', 'B1', 'C1']:
            hoja[celda].border = borde
            hoja[celda].fill = fondo_gris

        for i, fila in enumerate(dataframe_to_rows(tabla_final, index=False, header=False), start=2):
            hoja[f'A{i}'] = fila[0]
            hoja[f'B{i}'] = fila[1]
            hoja[f'C{i}'] = fila[2]

            hoja[f'A{i}'].border = borde
            hoja[f'B{i}'].border = borde
            hoja[f'C{i}'].border = borde

            # Si es la fila "Total general", aplicar fondo gris
            if fila[0] == "Total general":
                hoja[f'A{i}'].fill = fondo_gris
                hoja[f'B{i}'].fill = fondo_gris
                hoja[f'C{i}'].fill = fondo_gris

        # Llamar a la función graficas_barras para generar el gráfico de barras
        grafico_barras_path = graficas_barras(df, colores, nombre_hoja)

        # Insertar el gráfico de barras como imagen
        img_barras = Image(grafico_barras_path)
        hoja.add_image(img_barras, 'E5')  # Inserta la imagen en la celda E5

        # Llamar a la función graficas_pastel para generar el gráfico de pastel
        grafico_pastel_path = graficas_pastel(df, nombre_hoja)

        # Insertar el gráfico de pastel como imagen
        img_pastel = Image(grafico_pastel_path)
        hoja.add_image(img_pastel, 'E20')  # Inserta la imagen en la celda E20

    # Leer las hojas DTO y PCL
    df_dto = pd.read_excel(archivo_subido, sheet_name='DTO', parse_dates=['FECHA_VISADO'])
    df_pcl = pd.read_excel(archivo_subido, sheet_name='PCL', parse_dates=['FECHA_VISADO'])

    # Crear las hojas con las tablas y los gráficos (barras y pastel)
    crear_hoja("TABLA MES DTO", df_dto)
    crear_hoja("TABLA MES PCL", df_pcl)

    # Guardar el archivo con las hojas y gráficos
    output = BytesIO()
    libro.save(output)
    output.seek(0)
    return output



# ------------------------------------------------------------------------------ FLUJO ---------------------------------------------------------------------------------

def descargar_excel(output, nombre="archivo_procesado.xlsx"):
    st.download_button(
        label="📥 Descargar archivo",  
        data=output,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def procesar_archivos():
    archivo, tipo = subir_archivo()

    if archivo and tipo == "xlsx":
        output = generar_tablas_dto_y_pcl(archivo)
        descargar_excel(output, nombre="informe_dto_pcl.xlsx")
        st.success("✅ Archivo generado con éxito.")
    elif archivo and tipo == "csv":
        st.warning("Actualmente el procesamiento está disponible solo para archivos .xlsx con hojas DTO y PCL.")
