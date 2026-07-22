# Gestión de Convenios - CETPRO Cajamarca

Aplicación web desarrollada para la gestión, visualización y control de convenios institucionales del CETPRO Cajamarca. 

## Descripción del Proyecto
Este sistema permite centralizar la base de datos de convenios, facilitando la consulta de información clave como:
* Datos de empresas e instituciones.
* Representantes y contactos.
* Fechas de inicio, fin y días restantes de vigencia.
* Estado actual del convenio (Vigente / Vencido).

## Estructura de la Base de Datos
El sistema utiliza un archivo Excel (`BD_Convenios_CETPRO_CORREGIDA_FINAL_V1.xlsx`) que contiene la información centralizada en la hoja **BD_CONVENIOS**, incluyendo campos detallados desde ID de convenio hasta observaciones.

## Tecnologías Utilizadas
* **Python**: Lenguaje base del proyecto.
* **Streamlit**: Framework para la creación de la interfaz web interactiva.
* **Pandas**: Librería para el procesamiento y análisis de los datos del Excel.
* **Openpyxl**: Motor para la lectura de archivos de Excel.

## Uso
Para ejecutar la aplicación localmente:
1. Clonar el repositorio.
2. Instalar los requisitos: `pip install -r requirements.txt`.
3. Ejecutar el comando: `streamlit run PROYECTOS/app.py`.

---
*Desarrollado para la optimización de procesos administrativos en el CETPRO Cajamarca.*
