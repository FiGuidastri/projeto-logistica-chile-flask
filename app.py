import os
import io
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from flask import Flask, render_template, request, send_file, session, redirect, url_for

app = Flask(__name__)
# É crucial definir uma chave secreta para usar sessões no Flask
app.secret_key = 'super_secret_key_for_session_management' 

# =====================================================================================
#  TRANSLATIONS
# =====================================================================================
translations = {
    'en': {
        'title': "Holiday Rescheduler",
        'description': "This tool automates the rescheduling of deliveries. Upload your spreadsheet, enter the holiday day, and get the modified file.",
        'file_uploader_label': "1. Choose your scheduling spreadsheet",
        'number_input_label': "2. Enter the holiday day",
        'button_label': "Reschedule Spreadsheet",
        'processing_text': "Processing...",
        'report_header': "Operation Report",
        'success_message': "Your spreadsheet has been successfully rescheduled!",
        'download_button_label': "Download Rescheduled Spreadsheet",
        'download_file_suffix': "_rescheduled",
        'error_upload_file': "Please upload a spreadsheet to continue.",
        'error_no_file_part': "No file part in the request.",
        'error_no_file_selected': "No file selected.",
        # --- Logs ---
        'log_sheet_loaded': "Spreadsheet '{file_name}' loaded successfully.",
        'log_error_read_sheet': "ERROR: Could not read the spreadsheet. Please check if it is the correct file. Details: {error}",
        'log_error_day_not_found': "ERROR: The day {holiday_day} was not found in row 3 of the Delivery columns.",
        'log_holiday_identified': "Holiday identified in the Delivery column: {col_letter}",
        'log_warning_first_day': "Warning: The holiday is the first day of the period. It cannot be anticipated.",
        'log_rescheduled_with_substitution': "Delivery rescheduled (with substitution) from day {holiday_day} to column {col_letter}.",
        'log_rescheduling_complete': "Rescheduling completed. {tasks_moved} tasks were moved.",
    },
    'es': {
        'title': "Reprogramador de Feriados",
        'description': "Esta herramienta automatiza la reprogramación de entregas. Suba su planilla, ingrese el día feriado y obtenga el archivo modificado.",
        'file_uploader_label': "1. Elija su planilla de programación",
        'number_input_label': "2. Ingrese el día feriado",
        'button_label': "Reprogramar Planilla",
        'processing_text': "Procesando...",
        'report_header': "Reporte de Operación",
        'success_message': "¡Su planilla ha sido reprogramada exitosamente!",
        'download_button_label': "Descargar Planilla Reprogramada",
        'download_file_suffix': "_reprogramada",
        'error_upload_file': "Por favor, suba una planilla para continuar.",
        'error_no_file_part': "No hay archivo en la solicitud.",
        'error_no_file_selected': "No se ha seleccionado ningún archivo.",
        # --- Logs ---
        'log_sheet_loaded': "Planilla '{file_name}' cargada exitosamente.",
        'log_error_read_sheet': "ERROR: No se pudo leer la planilla. Por favor, verifique si es el archivo correcto. Detalles: {error}",
        'log_error_day_not_found': "ERROR: El día {holiday_day} no fue encontrado en la fila 3 de las columnas de Entrega.",
        'log_holiday_identified': "Feriado identificado en la columna de Entrega: {col_letter}",
        'log_warning_first_day': "Advertencia: El feriado es el primer día del período. No se puede anticipar.",
        'log_rescheduled_with_substitution': "Entrega reprogramada (con sustitución) del día {holiday_day} a la columna {col_letter}.",
        'log_rescheduling_complete': "Reprogramación completada. Se movieron {tasks_moved} tareas.",
    }
}

# =====================================================================================
#  BUSINESS LOGIC (Original function, unchanged)
# =====================================================================================
def process_spreadsheet(excel_file, holiday_day, texts):
    logs = []
    try:
        # A file vinda do Flask/request já pode ser lida diretamente
        workbook = openpyxl.load_workbook(excel_file)
        sheet_name = '01. Calendario SCL Abarrotes'
        sheet = workbook[sheet_name]
        logs.append(texts['log_sheet_loaded'].format(file_name=excel_file.filename))
    except Exception as e:
        logs.append(texts['log_error_read_sheet'].format(error=e))
        return None, logs

    delivery_columns = ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
    observations_column = 'CT'
    weekday_map = {'L': 1, 'M': 2, 'W': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 7}

    holiday_col_letter = None
    for col_letter in delivery_columns:
        day_in_sheet = sheet[f'{col_letter}3'].value
        if day_in_sheet == holiday_day:
            holiday_col_letter = col_letter
            break
    
    if not holiday_col_letter:
        logs.append(texts['log_error_day_not_found'].format(holiday_day=holiday_day))
        return None, logs
    
    logs.append(texts['log_holiday_identified'].format(col_letter=holiday_col_letter))

    holiday_col_index = column_index_from_string(holiday_col_letter)
    if holiday_col_index == column_index_from_string(delivery_columns[0]):
          logs.append(texts['log_warning_first_day'])
          return None, logs

    previous_col_index = holiday_col_index - 1
    previous_col_letter = get_column_letter(previous_col_index)
    
    tasks_moved = 0
    for row_index in range(8, sheet.max_row + 1):
        task_cell = sheet[f'{holiday_col_letter}{row_index}']
        if isinstance(task_cell.value, (int, float)) and 1 <= task_cell.value <= 6:
            destination_cell = sheet[f'{previous_col_letter}{row_index}']
            weekday_initial = sheet[f'{previous_col_letter}6'].value.upper()
            new_weekday_number = weekday_map.get(weekday_initial)
            
            if new_weekday_number:
                destination_cell.value = new_weekday_number
                task_cell.value = None
                log_message = texts['log_rescheduled_with_substitution'].format(
                    holiday_day=holiday_day,
                    col_letter=previous_col_letter
                )
                sheet[f'{observations_column}{row_index}'].value = log_message
                tasks_moved += 1
    
    logs.append(texts['log_rescheduling_complete'].format(tasks_moved=tasks_moved))
    
    return workbook, logs

# =====================================================================================
#  FLASK ROUTES
# =====================================================================================
@app.route('/', methods=['GET', 'POST'])
def index():
    lang = session.get('lang', 'en') # Default to English
    if 'lang' in request.args:
        lang = request.args.get('lang')
        if lang in translations:
            session['lang'] = lang
        return redirect(url_for('index'))

    texts = translations[lang]

    if request.method == 'POST':
        if 'spreadsheet' not in request.files:
            return render_template('index.html', texts=texts, error=texts['error_no_file_part'])
        
        file = request.files['spreadsheet']
        
        if file.filename == '':
            return render_template('index.html', texts=texts, error=texts['error_no_file_selected'])

        if file and file.filename.endswith('.xlsx'):
            holiday_day = int(request.form['holiday_day'])
            modified_workbook, logs = process_spreadsheet(file, holiday_day, texts)
            
            if modified_workbook:
                output = io.BytesIO()
                modified_workbook.save(output)
                output.seek(0)

                original_filename = file.filename
                base_name, extension = os.path.splitext(original_filename)
                new_filename = f"{base_name}{texts['download_file_suffix']}{extension}"

                # Store file data in session to be used by the download route
                session['file_data'] = output.getvalue()
                session['filename'] = new_filename
                
                return render_template('index.html', texts=texts, logs=logs, success=True)
            else:
                return render_template('index.html', texts=texts, logs=logs, success=False)

    # For GET request
    return render_template('index.html', texts=texts)

@app.route('/download')
def download_file():
    if 'file_data' in session and 'filename' in session:
        file_data = session.pop('file_data', None)
        filename = session.pop('filename', None)
        
        return send_file(
            io.BytesIO(file_data),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)