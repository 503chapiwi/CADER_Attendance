Reportes MAGA Totonicapán
Aplicación Streamlit que convierte las asistencias del formulario Kobo "Planillas MAGA Toto" en los dos reportes que las agencias municipales entregan cada mes:

Listado de Beneficiarios (REGISTRO MENSUAL DE BENEFICIARIOS POR INTERVENCIÓN, hojas HOGAR RURAL / AGRÍCOLA / PECUARIO)
Informe Mensual de Capacitaciones (FORMATO DE CAPACITACIONES, con la sección del municipio elegido llenada y la hoja RESUMEN funcionando con sus fórmulas originales)
Cómo funciona
Descarga las boletas enviadas al formulario Kobo (uid aMAY4yQjPMm3Nvw6w3AdEv en kf.kobotoolbox.org) usando el API.
Descarga automáticamente la base de participantes CADER (participantes_toto_2026.csv, archivo multimedia del propio formulario) para completar los datos de cada CUI: nombres, sexo, año de nacimiento, pueblo, comunidad lingüística, discapacidad y retornado. También usa toto_promoters.csv para poner el nombre del promotor según el CADER.
Las personas registradas como nuevos integrantes en las boletas también entran al listado, con los datos capturados en el formulario.
CUIs que no aparecen en ninguna fuente salen en el listado solo con el CUI y se muestran como advertencia en la app.
Reglas aplicadas (acordadas):

En el listado, cada persona aparece una sola vez por hoja al mes; si tuvo varias capacitaciones de la misma categoría, se juntan en la columna de descripción. "Valor del beneficio Q" queda vacío.
En el informe, los conteos H/M son los números digitados en Kobo (promotores + otros participantes); los totales los calculan las fórmulas de la plantilla.
El informe sale con la plantilla departamental completa y solo la sección del municipio elegido llenada. Máximo 10 eventos por sección (limitación de la plantilla; la app avisa si hay más).
Correr localmente
pip install -r requirements.txt
streamlit run app.py
El token de Kobo se busca en este orden:

st.secrets["KOBO_TOKEN"] (archivo .streamlit/secrets.toml, ver secrets.toml.example)
kobo_token.txt junto a app.py
Campo de texto en la app
Publicar en Streamlit Community Cloud
Suba esta carpeta a un repositorio de GitHub — sin kobo_token.txt (el token da acceso total a su cuenta Kobo; si el repo es público, cualquiera podría verlo).
En https://share.streamlit.io cree la app apuntando a app.py.
En Settings → Secrets agregue: KOBO_TOKEN = "su_token_aqui"
Comparta el enlace con los extensionistas. Ellos solo eligen municipio, mes y reportes, y descargan los dos archivos.
Archivos
app.py — interfaz Streamlit
kobo_client.py — conexión con el API de Kobo y traducción código→etiqueta
reports.py — lógica de llenado de las dos plantillas
plantillas/plantilla_listado.xlsx — plantilla del listado (original MAGA)
plantillas/plantilla_informe.xlsx — plantilla del informe (original MAGA, con las filas de ejemplo limpiadas y las fórmulas U/V restauradas)
Mantenimiento
Si cambian la base de participantes, solo hay que reemplazar participantes_toto_2026.csv en los archivos multimedia del formulario en Kobo — la app siempre baja la versión actual (o se puede subir un archivo distinto en "Opciones avanzadas").
Si agregan preguntas al formulario, los nombres de campo usados están en reports.py (generar_listado / generar_informe).
El botón "Actualizar datos de Kobo" limpia el caché (que dura 10 minutos).
