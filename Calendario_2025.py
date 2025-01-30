#Se instala bibliotecas:  pandas, os , xlsxwriter
import calendar
import pandas as pd
import os
#aqui se debe especificar la ruta deseada de acuerdo a su pc"
ruta_deseada = "C:/Users/Nicolas Cruz/Documents/Marketing/mrt/data"
os.makedirs(ruta_deseada, exist_ok=True)
# Crear un diccionario de DataFrames con formato de calendario
sheets_calendar = {}
months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
for month_num, month_name in enumerate(months, start=1):
    # Obtener el primer día de la semana y el número de días del mes
    first_weekday, num_days = calendar.monthrange(2025, month_num)

    # Crear una matriz de calendario (6 semanas x 7 días)
    cal_matrix = [["" for _ in range(7)] for _ in range(6)]
    
    # Llenar la matriz con las fechas del mes
    day = 1
    for row in range(6):
        for col in range(7):
            if (row == 0 and col < first_weekday) or day > num_days:
                cal_matrix[row][col] = ""  # Espacios vacíos antes y después del mes
            else:
                # Crear el cuadro con la estructura de contenido
                cal_matrix[row][col] = f"{day}\n---\nTipo de Contenido:\nPlataforma:\nTexto/Copia:\nImagen/Video:\nHashtags:\nHora:\nEncargados:"
                day += 1
    
    # Convertir la matriz en DataFrame
    sheets_calendar[month_name] = pd.DataFrame(cal_matrix, columns=["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"])

# Guardar el archivo Excel con formato de calendario
excel_calendar_path = "C:/Users/Nicolas Cruz/Documents/Marketing/mrt/data/calendario_2025.xlsx"
with pd.ExcelWriter(excel_calendar_path, engine="xlsxwriter") as writer:
    for month, df in sheets_calendar.items():
        df.to_excel(writer, sheet_name=month, index=False)

excel_calendar_path
