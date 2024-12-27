from datetime import timedelta, date
import calendar
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Autenticación y acceso a la hoja de cálculo de Google
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('cajas-cpcc-0648c6abbd90.json', scope)
client = gspread.authorize(creds)

# Abre la hoja de cálculo por su nombre
spreadsheet = client.open("Cajas CPCC 2025-01")

# Configuración inicial
first_day = date(2025, 1, 1)
_, num_days = calendar.monthrange(first_day.year, first_day.month)

month = 'Enero'
spanish_days = {
    'Sun': 'Domingo',
    'Mon': 'Lunes',
    'Tue': 'Martes',
    'Wed': 'Miércoles',
    'Thu': 'Jueves',
    'Fri': 'Viernes',
    'Sat': 'Sábado'
}

# Listado de operaciones masivas
requests = []
template_sheet_id = spreadsheet.get_worksheet(0).id  # Obtén el ID de la primera pestaña como plantilla

# Crear operaciones de duplicación y actualización
for i in range(1, num_days + 1):
    day = first_day + timedelta(days=i - 1)
    day_name = spanish_days.get(day.strftime("%a"))  # Nombre corto del día (Mon, Tue, etc.)
    sheet_title = f'{day.day} {day_name[:3]}'

    # Crear la pestaña basada en la plantilla
    requests.append({
        "duplicateSheet": {
            "sourceSheetId": template_sheet_id,
            "insertSheetIndex": i,
            "newSheetName": sheet_title
        }
    })

# Ejecutar todas las duplicaciones en una sola operación
spreadsheet.batch_update({"requests": requests})

# Actualizar las celdas en las nuevas pestañas
for i in range(1, num_days + 1):
    day = first_day + timedelta(days=i - 1)
    day_name = spanish_days.get(day.strftime("%a"))  # Nombre corto del día (Mon, Tue, etc.)
    day_num = day.day

    sheet_title = f'{day.day} {day_name[:3]}'
    sheet = spreadsheet.worksheet(sheet_title)
    sheet.update(range_name='B5', values=[[f'{day_name} {day.day} de {month}']])


print(f"Se han creado y actualizado todas las pestañas para {num_days} días.")
