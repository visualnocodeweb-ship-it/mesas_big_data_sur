
import pandas as pd

# Cargar el archivo de Excel
df = pd.read_excel("usuarios_2025-10-25.xlsx")

# Buscar la fila de Jose Fabian D'ALOI y actualizar la Localidad
df.loc[df['Nombre Completo'] == "Jose Fabian D'ALOI", 'Localidad'] = "JUNIN DE LOS ANDES"

# Guardar el DataFrame actualizado en el archivo de Excel
df.to_excel("usuarios_2025-10-25.xlsx", index=False)

print("Archivo de Excel actualizado correctamente.")
