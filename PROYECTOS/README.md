# Buscador de Convenios - CETPRO Cajamarca

Aplicación local (sin conexión a internet) para buscar convenios por
**DNI del representante**, **RUC**, empresa, representante o programa formativo.

## 📁 Estructura del repositorio en GitHub

Sube estos 4 archivos juntos, en la raíz del repositorio:

```
tu-repositorio/
├── app.py
├── requirements.txt
├── logo_cetpro.png                              ← tu logo (opcional, mismo nombre exacto)
└── BD_Convenios_CETPRO_CORREGIDA_FINAL_V1.xlsx  ← tu base de datos
```

> Si tu logo tiene otro nombre, cambia la línea `LOGO_PATH = "logo_cetpro.png"`
> en `app.py` (sección 1, CONFIGURACIÓN) por el nombre real del archivo.
> Si no subes logo, la app funciona igual, solo no muestra la imagen.

## 🔑 Acceso

- **Usuario:** `CETPRO CAJAMARCA`
- **Contraseña:** `cetprocajamarca#2026`

Puedes cambiarlos en `app.py`, sección 1 (`USUARIO_VALIDO` y `CONTRASENA_VALIDA`).

## ▶️ Cómo ejecutarlo

### Opción A: Streamlit Community Cloud (gratis, recomendado)
1. Sube los 4 archivos a un repositorio de GitHub.
2. Entra a https://share.streamlit.io con tu cuenta de GitHub.
3. Clic en "New app", elige el repositorio y como archivo principal: `app.py`.
4. Deploy. En 1-2 minutos tendrás tu link público.

### Opción B: En tu propia computadora
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 🔒 Privacidad
La aplicación **no se conecta a internet**: toda la búsqueda ocurre
localmente sobre el archivo Excel cargado en el mismo repositorio.

## 🔎 Cómo buscar
Escribe en el cuadro de búsqueda cualquiera de estos datos:
- DNI del representante (búsqueda exacta)
- RUC (búsqueda exacta)
- Nombre de la empresa/institución (búsqueda parcial)
- Nombre del representante (búsqueda parcial)
- Programa formativo (búsqueda parcial)

Los resultados se muestran en formato de párrafo vertical, con íconos,
uno debajo del otro (una tarjeta por convenio encontrado).
