
import pandas as pd

# Cargar el archivo de Excel
df = pd.read_excel("usuarios_2025-10-25.xlsx")

# Nuevos datos
new_data = {
    "Nombre Completo": "Adriana lorena Thuman",
    "Afinidades": 181 + 5 + 0,
    "Localidad": "SAN MARTIN DE LOS ANDES",
    "Mesa": ""  # Assuming no mesa is provided
}

# Crear un DataFrame para los nuevos datos
new_df = pd.DataFrame([new_data])

# Añadir los nuevos datos al DataFrame existente
df = pd.concat([df, new_df], ignore_index=True)

# Guardar el DataFrame actualizado en el archivo de Excel
df.to_excel("usuarios_2025-10-25.xlsx", index=False)

print("Archivo de Excel actualizado correctamente.")
