import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Border, Side, PatternFill

# -------------------------- FUNCIONES DE PROCESAMIENTO Y GENERACIÓN DE TABLAS ---------------------------
def generar_tablas_estado_informe(archivo, tipo):
    # Cargar el archivo (CSV o Excel)
    df_base = cargar_archivo(archivo, tipo)
    
    if df_base is not None:
        # Verificar que las columnas necesarias existan en df_base
        if "ESTADO_INFORME" in df_base.columns and "NOTIFICADOR" in df_base.columns:
            # Agrupar los datos por 'ESTADO_INFORME' y 'NOTIFICADOR' y contar los registros
            conteo = df_base.groupby(['ESTADO_INFORME', 'NOTIFICADOR']).size().unstack(fill_value=0)

            # Calcular TOTAL GENERAL como la suma de las filas (es decir, sumar los valores de cada estado)
            conteo['TOTAL GENERAL'] = conteo.sum(axis=1)

            # Calcular TOTAL % para cada valor respecto al TOTAL GENERAL (como porcentaje)
            total_general_sum = conteo['TOTAL GENERAL'].sum()  # Sumar los TOTAL GENERAL para calcular el porcentaje global

            # Crear una nueva columna 'TOTAL %' que muestra el porcentaje de cada celda respecto al TOTAL GENERAL
            conteo['TOTAL %'] = (conteo['TOTAL GENERAL'] / total_general_sum) * 100

            # Añadir el símbolo de porcentaje al final de 'TOTAL %'
            conteo['TOTAL %'] = conteo['TOTAL %'].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else '')

            # Crear un archivo Excel
            libro = Workbook()

            # Hoja para la tabla procesada
            hoja_procesada = libro.active
            hoja_procesada.title = "Tabla Procesada"[:31]  # Asegurar que el nombre no exceda 31 caracteres
            
            # Escribir los encabezados
            hoja_procesada.cell(row=1, column=1, value="ESTADO INFORME")
            for col_idx, notificador in enumerate(conteo.columns[:-2], start=2):  # Ignoramos 'TOTAL GENERAL' y 'TOTAL %'
                hoja_procesada.cell(row=1, column=col_idx, value=notificador)

            hoja_procesada.cell(row=1, column=len(conteo.columns)-1, value="TOTAL GENERAL")
            hoja_procesada.cell(row=1, column=len(conteo.columns), value="TOTAL %")

            # Escribir los datos de la tabla procesada
            row_start = 2
            for estado, valores in conteo.iterrows():
                hoja_procesada.cell(row=row_start, column=1, value=estado)  # ESTADO_INFORME
                # Escribir los valores de NOTIFICADOR
                for col_idx, notificador in enumerate(conteo.columns[:-2], start=2):  # Ignoramos 'TOTAL GENERAL' y 'TOTAL %'
                    hoja_procesada.cell(row=row_start, column=col_idx, value=valores.get(notificador, 0))
                hoja_procesada.cell(row=row_start, column=len(valores)-2, value=valores.get('TOTAL GENERAL', 0))  # TOTAL GENERAL
                hoja_procesada.cell(row=row_start, column=len(valores)-1, value=valores.get('TOTAL %', ''))  # TOTAL %
                row_start += 1

            # Añadir la fila 'TOTAL GENERAL' al final (última fila de la tabla)
            hoja_procesada.cell(row=row_start, column=1, value="TOTAL GENERAL")  # Coloca el texto 'TOTAL GENERAL'
            for col_idx in range(2, len(conteo.columns)-1):  # Excluimos 'TOTAL GENERAL' y 'TOTAL %'
                hoja_procesada.cell(row=row_start, column=col_idx, value=conteo.iloc[:, col_idx - 1].sum())  # Suma de las columnas

            hoja_procesada.cell(row=row_start, column=len(conteo.columns)-1, value='')  # Dejar en blanco la celda de TOTAL % para la fila 'TOTAL GENERAL'

            # Crear la hoja BASE con la unión de DTO y PCL
            hoja_base = libro.create_sheet("BASE"[:31])  # Asegurar que el nombre no exceda 31 caracteres
            for r_idx, row in enumerate(dataframe_to_rows(df_base, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    hoja_base.cell(row=r_idx, column=c_idx, value=value)

            # Aplicar formato a la tabla procesada
            for col in hoja_procesada.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column name
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                hoja_procesada.column_dimensions[column].width = adjusted_width

            # Crear los bordes y fondo gris para los encabezados y la fila 'TOTAL GENERAL'
            borde = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000")
            )
            fondo_gris = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            fondo_gris_total = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")  # Gris más oscuro para 'TOTAL GENERAL'

            # Estilo para el encabezado
            for cell in hoja_procesada["1:1"]:
                cell.border = borde
                cell.fill = fondo_gris

            # Estilo para la fila 'TOTAL GENERAL'
            for cell in hoja_procesada[row_start]:
                cell.fill = fondo_gris_total
                cell.border = borde

            # Aplicar bordes a todas las celdas de la tabla procesada
            for row in hoja_procesada.iter_rows(min_row=2, max_row=row_start, min_col=1, max_col=len(conteo.columns)):
                for cell in row:
                    cell.border = borde

            # Guardar el archivo generado
            output = BytesIO()
            libro.save(output)
            output.seek(0)
            output.flush()  # Asegura que los datos se escriban correctamente
            return output
        else:
            st.error("El archivo no contiene las columnas necesarias: 'ESTADO_INFORME' y 'NOTIFICADOR'.")
            return None


# ------------------------ FUNCIONES DE SUBIDA Y DESCARGA -------------------------------

# Función para descargar el archivo generado
def descargar_excel(output, nombre="archivo_procesado.xlsx"):
    st.download_button(
        label="📥 Descargar archivo",  
        data=output,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Función para subir el archivo
def subir_archivo2():
    archivo = st.file_uploader("Sube un archivo (.xlsx o .csv)", type=["xlsx", "csv"], key="file_uploader")
    
    if archivo is not None:
        try:
            nombre_archivo = archivo.name.lower()

            if nombre_archivo.endswith(".xlsx"):
                st.success("¡Archivo Excel válido!")
                return archivo, "xlsx"
            elif nombre_archivo.endswith(".csv"):
                st.success("¡Archivo CSV válido!")
                return archivo, "csv"
            else:
                st.warning("El archivo debe ser de tipo .xlsx o .csv")
                return None, None

        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
            return None, None

    return None, None

# Función para cargar y limpiar el archivo (maneja CSV y Excel)
def cargar_archivo(archivo, tipo):
    try:
        if tipo == "xlsx":
            # Cargar ambas hojas (DTO y PCL)
            df_dto = pd.read_excel(archivo, sheet_name='DTO')
            df_pcl = pd.read_excel(archivo, sheet_name='PCL')

            # Unir ambas hojas en un solo DataFrame
            df_base = pd.concat([df_dto, df_pcl], ignore_index=True)

        elif tipo == "csv":
            df_base = pd.read_csv(archivo, on_bad_lines='skip', delimiter=",")  # 'skip' ignora las líneas mal formadas

        # Limpiar posibles filas con datos inconsistentes
        df_base.dropna(how='all', inplace=True)  # Eliminar filas vacías
        df_base = df_base.reset_index(drop=True)  # Resetear el índice

        return df_base
    except Exception as e:
        st.error(f"Error al procesar el archivo {tipo}: {e}")
        return None

# ---------------------------- FLUJO  --------------------------

# Función para procesar el archivo y generar la tabla
def procesar_archivos2():
    archivo, tipo = subir_archivo2()

    if archivo and tipo in ["xlsx", "csv"]:
        output = generar_tablas_estado_informe(archivo, tipo)
        if output:
            descargar_excel(output, nombre="informe_estado_informe.xlsx")
            st.success("✅ Archivo generado con éxito.")
    else:
        st.error("No se ha cargado un archivo válido.")
