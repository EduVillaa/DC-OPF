import pandas as pd
from pathlib import Path

ruta = Path(r"/Users/edu/Desktop/comparacion.xlsx")
ruta_salida = Path(r"/Users/edu/Desktop/comparacion_resultados.xlsx")

df1 = pd.read_excel(ruta, sheet_name="Sheet1", header=1, usecols=[0, 1])
df2 = pd.read_excel(ruta, sheet_name="Sheet2", header=1, usecols=[0, 1])

# Palabras comunes
commonMots = df1.loc[df1["Mot"].isin(df2["Mot"]), ["Mot"]].copy()

same_quantity_rows = []
different_quantity_rows = []

for mot in commonMots["Mot"]:
    row_pos_df1 = df1[df1["Mot"] == mot].index[0]
    row_pos_df2 = df2[df2["Mot"] == mot].index[0]

    cantidad_df1 = df1.loc[row_pos_df1, "cantidad"]
    cantidad_df2 = df2.loc[row_pos_df2, "cantidad"]

    if cantidad_df1 == cantidad_df2:
        same_quantity_rows.append({"Mot": mot})
    else:
        different_quantity_rows.append({
            "Mot": mot,
            "cantidad df1": cantidad_df1,
            "cantidad df2": cantidad_df2
        })

commonMots_with_same_quantity = pd.DataFrame(same_quantity_rows)
commonMots_with_different_quantity = pd.DataFrame(different_quantity_rows)

# Palabras únicas de cada df
unique_df1_Mots = df1.loc[~df1["Mot"].isin(df2["Mot"]), :].copy()
unique_df2_Mots = df2.loc[~df2["Mot"].isin(df1["Mot"]), :].copy()

print(commonMots_with_same_quantity)
print("\n")
print(commonMots_with_different_quantity)
print("\n")
print(unique_df1_Mots)
print("\n")
print(unique_df2_Mots)

# Exportar todos los DataFrames a una misma hoja de Excel
with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
    sheet_name = "Resultados"
    current_row = 0

    def write_df_with_title(df, title, writer, sheet_name, startrow):
        # Escribir título
        pd.DataFrame([[title]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            index=False,
            header=False
        )

        # Escribir dataframe debajo del título
        df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow + 1,
            index=False
        )

        # Devolver siguiente fila libre, dejando 2 filas en blanco
        return startrow + len(df) + 4

    current_row = write_df_with_title(
        commonMots_with_same_quantity,
        "Palabras comunes con la misma cantidad",
        writer,
        sheet_name,
        current_row
    )

    current_row = write_df_with_title(
        commonMots_with_different_quantity,
        "Palabras comunes con distinta cantidad",
        writer,
        sheet_name,
        current_row
    )

    current_row = write_df_with_title(
        unique_df1_Mots,
        "Palabras unicas de Sheet1",
        writer,
        sheet_name,
        current_row
    )

    current_row = write_df_with_title(
        unique_df2_Mots,
        "Palabras unicas de Sheet2",
        writer,
        sheet_name,
        current_row
    )

print(f"Archivo exportado en: {ruta_salida}")
