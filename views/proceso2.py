import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side, PatternFill
import csv

# Función para procesar y generar la tabla de ESTADO_INFORME y NOTIFICADORES
def generar_tabla_estado_notificadores(df):
    # Agrupar los datos por 'ESTADO_INFORME' y 'NOTIFICADOR' y contar los registros
    conteo = df.groupby(['ESTADO_INFORME', 'NOTIFICADOR']).size().unstack(fill_value=0)

    # Agregar columna 'TOTAL' para cada 'ESTADO_INFORME'
    conteo['TOTAL'] = conteo.sum(axis=1)

    # Agregar fila 'TOTAL GENERAL'
    total_general = conteo.sum().sum()
    conteo.loc['TOTAL GENERAL'] = conteo.sum()

    # Calcular el porcentaje por cada celda respecto al total general
    conteo_percentage = conteo.div(total_general).multiply(100).round(2)
    conteo_percentage['TOTAL'] = conteo_percentage.sum(axis=1)  # Asegurar que la fila 'TOTAL' sea el 100%

    # Crear la tabla con el formato deseado
    tabla_final = conteo.join(conteo_percentage, lsuffix='_TOTAL', rsuffix='_PCT')

    return tabla_final


# Función para generar el archivo Excel con la tabla y los gráficos
def generar_tablas_estado_informe(archivo_subido):
    archivo_subido.seek(0)
    df = pd.read_excel(archivo_subido)

    # Verificar que las columnas necesarias existan
    if "ESTADO_INFORME" in df.columns and "NOTIFICADOR" in df.columns:
        # Generar la tabla de ESTADO_INFORME y NOTIFICADOR
        tabla_final = generar_tabla_estado_notificadores(df)

        # Crear un archivo Excel
        libro = Workbook()
        hoja = libro.active
        hoja.title = "Estado Informe y Notificadores"

        # Escribir los encabezados de la tabla
        for col_num, col_name in enumerate(tabla_final.columns, 1):
            hoja.cell(row=1, column=col_num, value=col_name)

        # Escribir los datos de la tabla
        for row_num, row in enumerate(dataframe_to_rows(tabla_final, index=True, header=False), 2):
            for col_num, cell_value in enumerate(row, 1):
                hoja.cell(row=row_num, column=col_num, value=cell_value)

        # Aplicar formato a la tabla
        for col in hoja.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            hoja.column_dimensions[column].width = adjusted_width

        # Agregar bordes y fondo gris para los encabezados
        borde = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000")
        )
        fondo_gris = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        for cell in hoja["1:1"]:
            cell.border = borde
            cell.fill = fondo_gris

        # Guardar el archivo generado
        output = BytesIO()
        libro.save(output)
        output.seek(0)
        return output
    else:
        st.error("El archivo no contiene las columnas necesarias: 'ESTADO_INFORME' y 'NOTIFICADOR'.")
        return None


# Función para subir el archivo
def subir_archivo2():
    archivo = st.file_uploader("Sube un archivo (.xlsx o .csv)", type=["xlsx", "csv"], key="file_uploader")
    
    if archivo is not None:
        try:
            nombre_archivo = archivo.name.lower()

            if nombre_archivo.endswith(".xlsx"):
                # Verificar que el archivo Excel tenga las hojas necesarias
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
                # Intentar leer el archivo CSV con opciones avanzadas para manejar líneas problemáticas
                try:
                    # Leer el archivo CSV con un manejo más robusto de las líneas problemáticas
                    df = pd.read_csv(archivo, 
                                     on_bad_lines='skip',  # Omitir líneas con problemas
                                     quoting=csv.QUOTE_NONE, # No esperar comillas en los campos
                                     delimiter=',',          # Especificar el delimitador
                                     engine='python')       # Usar el motor de Python, más flexible

                    if "DTO" in df.columns and "PCL" in df.columns:
                        st.success("¡Archivo CSV válido! Se encontraron las columnas DTO y PCL.")
                        return archivo, "csv"
                    else:
                        st.warning("El archivo CSV no contiene las columnas 'DTO' y 'PCL'.")
                        return None, None
                except Exception as e:
                    st.error(f"Error al leer el archivo CSV: {e}")
                    return None, None

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            return None, None

    return None, None


# Función para descargar el archivo generado
def descargar_excel(output, nombre="archivo_procesado.xlsx"):
    st.download_button(
        label="📥 Descargar archivo",  
        data=output,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# Función para procesar el archivo y generar la tabla
def procesar_archivos2():
    archivo, tipo = subir_archivo2()

    if archivo and tipo == "xlsx":
        output = generar_tablas_estado_informe(archivo)
        if output:
            descargar_excel(output, nombre="informe_estado_informe.xlsx")
            st.success("✅ Archivo generado con éxito.")
    elif archivo and tipo == "csv":
        st.warning("Actualmente el procesamiento está disponible solo para archivos .xlsx con las columnas 'ESTADO_INFORME' y 'NOTIFICADOR'.")
    else:
        st.error("No se ha cargado un archivo válido.")
