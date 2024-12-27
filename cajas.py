import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import timedelta, date
import calendar
import time

# Autenticación y acceso a la hoja de cálculo de Google
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('cajas-cpcc-0648c6abbd90.json', scope)
client = gspread.authorize(creds)

# Abre la hoja de cálculo por su nombre
spreadsheet = client.open("Cajas CPCC 2025-02")
# spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1mat3ZRG3deLubxz8ypW_iSoy9Jo0tGwvX-zY8I1j4nw/edit#gid=0")

# Selecciona la primera pestaña
first_sheet = spreadsheet.get_worksheet(0)

first_day = date(2025,2,1)

_, num_days = calendar.monthrange(first_day.year, first_day.month)

num_days= 28
month='Febrero'
spanish_days={
    'Sun' : 'Domingo',
    'Mon' : 'Lunes',
    'Tue' : 'Martes',
    'Wed' : 'Miércoles',
    'Thu' : 'Jueves',
    'Fri' : 'Viernes',
    'Sat' : 'Sábado'
}
# Crear pestañas para cada día del mes
for i in range(num_days,0,-1):
    day = first_day + timedelta(days=i)
    day_name = spanish_days.get(day.strftime("%a"))  # Nombre corto del día (Mon, Tue, etc.)
    day_num = day.day

    # Nombra las pestañas con el formato '1 Mar', '2 Mie', etc.
    sheet_title = f'{day.day}  {day_name[:3]}'
    
    # Duplica la primera pestaña y le cambia el nombre
    new_sheet = first_sheet.duplicate()
    new_sheet.update_title(sheet_title)
    new_sheet.update_acell('B5',f'{day_name} {day.day} de {month}')
    
    print(f"Pestaña '{sheet_title}' creada.")
    time.sleep(1)
