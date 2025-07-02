import streamlit as st
from views.proceso1 import procesar_archivos
from views.proceso2 import procesar_archivos2

# Título
st.title("🔔 Notificaciones")

# Menú 
opciones_menu = ["Proceso 1", "Proceso 2"]

# Mostrar el menú en la barra lateral
opcion_seleccionada = st.sidebar.selectbox("Seleccione un proceso", opciones_menu)

# ------------------------------------------------------------------------------ Proceso 1 ---------------------------------------------------------------------------------
if opcion_seleccionada == "Proceso 1":
    st.subheader("Graficación año DTO y PCL")
    procesar_archivos()  
# ------------------------------------------------------------------------------ Proceso 2 ---------------------------------------------------------------------------------
elif opcion_seleccionada == "Proceso 2":
    st.subheader("Graficación Medicina Laboral")
    procesar_archivos2()
else:
    st.write("Por favor, selecciona un proceso del menú.")



