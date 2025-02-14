import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# 1️⃣ 📂 Cargar el archivo Excel sin modificar los datos
archivo_excel = "./datos_limpios.xlsx.xlsx"  # Nombre del archivo existente
hoja_datos = "Resumen de Ventas"  # Nombre de la hoja donde están los datos

# Cargar el libro de Excel
wb = load_workbook(archivo_excel)
ws = wb[hoja_datos]

# 2️⃣ 📊 Leer los datos desde la hoja existente (sin modificar)
df = pd.read_excel(archivo_excel, sheet_name=hoja_datos)

# Convertir la columna de fecha si es necesario
if "Fecha" in df.columns:
    df["Fecha"] = pd.to_datetime(df["Fecha"])

# 3️⃣ 🔄 Crear una tabla pivote si no está en el Excel
pivot_df = df.pivot_table(values='Ventas', index='Fecha', columns='Categoría', aggfunc='sum', fill_value=0)

# 4️⃣ 📈 Generar gráficos

# Gráfico de líneas
plt.figure(figsize=(8, 4))
sns.lineplot(data=pivot_df, marker='o')
plt.title('Ventas por Categoría a lo Largo del Tiempo')
plt.xlabel('Fecha')
plt.ylabel('Ventas')
plt.xticks(rotation=45)
plt.legend(title='Categoría')
plt.tight_layout()
plt.savefig("grafico_actualizado.png")  # Guardar imagen

# 5️⃣ 📂 Insertar el gráfico en el archivo Excel sin borrar datos

# Cargar la imagen y añadirla a la hoja de Excel
img = Image("grafico_actualizado.png")
ws.add_image(img, "E5")  # Ubicación en la hoja

# Guardar el archivo sin perder la información previa
wb.save(archivo_excel)

print(f"✅ Gráficos actualizados en: {archivo_excel}")
