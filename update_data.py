
import pandas as pd

# Cargar el archivo de Excel
try:
    df = pd.read_excel("usuarios_2025-10-25.xlsx")
except FileNotFoundError:
    # Si el archivo no existe, crea un DataFrame vacío con las columnas necesarias
    df = pd.DataFrame(columns=['Nombre Completo', 'Afinidades', 'Localidad', 'Mesa'])

# Nuevos datos
new_data = {
    'Nombre Completo': "Jose Fabian D'ALOI",
    'Afinidades': 130 + 16 + 27,
    'Localidad': "Junin de los Andes",
    'Mesa': 133
}

# Crear un DataFrame para los nuevos datos
new_df = pd.DataFrame([new_data])

# Añadir los nuevos datos al DataFrame existente
df = pd.concat([df, new_df], ignore_index=True)


# Guardar el DataFrame actualizado en el archivo de Excel
df.to_excel("usuarios_2025-10-25.xlsx", index=False)

print("Archivo de Excel actualizado correctamente.")
