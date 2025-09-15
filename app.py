import os
import io
import uuid
import openpyxl
import xlwings as xw
from flask import Flask, render_template, request, send_from_directory, session, redirect, url_for
from collections import Counter

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_session_management'

TMP_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tmp')
if not os.path.exists(TMP_FOLDER):
    os.makedirs(TMP_FOLDER)

# (A seção de 'translations' continua a mesma, foi omitida por brevidade)
translations = {
    'en': {
        'title': "🤖 Automatic Holiday Rescheduler", 'description': "This tool automates the rescheduling of deliveries. Upload your spreadsheet, enter the holiday day, and get the modified file.",
        'file_uploader_label': "1. Choose your scheduling spreadsheet", 'number_input_label': "2. Enter the holiday day", 'button_label': "Reschedule Spreadsheet",
        'processing_text': "Processing...", 'report_header': "Operation Report", 'success_message': "Your spreadsheet has been successfully rescheduled!",
        'download_button_label': "Download Rescheduled Spreadsheet", 'view_report_button_label': "View Spreadsheet Report",
        'report_title': "Rescheduling Report", 'total_deliveries': "Total Deliveries", 'rescheduled_deliveries': "Rescheduled Deliveries", 'total_carriers': "Affected Carriers", 'total_chains': "Affected Chains",
        'deliveries_by_carrier': "Rescheduled by Carrier", 'deliveries_by_chain': "Rescheduled by Chain", 'deliveries_by_day': "Deliveries by Day", 'rescheduled_stores_list': "List of Rescheduled Stores",
        'back_to_home': "Back to Home", 'close_button': "Close", 'download_file_suffix': "_rescheduled", 'error_upload_file': "Please upload a spreadsheet to continue.",
        'error_no_file_part': "No file part in the request.", 'error_no_file_selected': "No file selected.",
        'error_wrong_file_type': "Invalid file type. Please upload a .xlsx file.",
        'log_sheet_loaded': "Spreadsheet '{file_name}' loaded successfully.", 'log_error_read_sheet': "ERROR: Could not read the spreadsheet. Please check if it is the correct file. Details: {error}",
        'log_error_day_not_found': "ERROR: The day {holiday_day} was not found in row 3 of the Delivery columns.", 'log_holiday_identified': "Holiday identified in the Delivery column: {col_letter}",
        'log_warning_first_day': "Warning: The holiday is the first day of the period. It cannot be anticipated.",
        'log_rescheduled_with_substitution': "Delivery rescheduled (with substitution) from day {holiday_day} to column {col_letter}.", 'log_rescheduling_complete': "Rescheduling completed. {tasks_moved} tasks were moved.",
    },
    'es': {
        'title': "🤖 Reprogramador Automático de Feriados", 'description': "Esta herramienta automatiza la reprogramación de entregas. Suba su planilla, ingrese el día feriado y obtenga el archivo modificado.",
        'file_uploader_label': "1. Elija su planilla de programación", 'number_input_label': "2. Ingrese el día feriado", 'button_label': "Reprogramar Planilla",
        'processing_text': "Procesando...", 'report_header': "Reporte de Operación", 'success_message': "¡Su planilla ha sido reprogramada exitosamente!",
        'download_button_label': "Descargar Planilla Reprogramada", 'view_report_button_label': "Ver Reporte de la Planilla",
        'report_title': "Reporte de Reprogramación", 'total_deliveries': "Total de Entregas", 'rescheduled_deliveries': "Entregas Reprogramadas", 'total_carriers': "Transportistas Afectados", 'total_chains': "Cadenas Afectadas",
        'deliveries_by_carrier': "Reprogramadas por Transportista", 'deliveries_by_chain': "Reprogramadas por Cadena", 'deliveries_by_day': "Entregas por Día", 'rescheduled_stores_list': "Lista de Tiendas Reprogramadas",
        'back_to_home': "Volver al Inicio", 'close_button': "Cerrar", 'download_file_suffix': "_reprogramada", 'error_upload_file': "Por favor, suba una planilla para continuar.",
        'error_no_file_part': "No hay archivo en la solicitud.", 'error_no_file_selected': "No se ha seleccionado ningún archivo.",
        'error_wrong_file_type': "Tipo de archivo inválido. Por favor, suba un archivo .xlsx.",
        'log_sheet_loaded': "Planilla '{file_name}' cargada exitosamente.", 'log_error_read_sheet': "ERROR: No se pudo leer la planilla. Por favor, verifique si es el archivo correcto. Detalles: {error}",
        'log_error_day_not_found': "ERROR: El día {holiday_day} no fue encontrado en la fila 3 de las columnas de Entrega.", 'log_holiday_identified': "Feriado identificado en la columna de Entrega: {col_letter}",
        'log_warning_first_day': "Advertencia: El feriado es el primer día del período. No se puede anticipar.",
        'log_rescheduled_with_substitution': "Entrega reprogramada (con sustitución) del día {holiday_day} a la columna {col_letter}.", 'log_rescheduling_complete': "Reprogramación completada. Se movieron {tasks_moved} tareas.",
    }
}
# =====================================================================================
#  BUSINESS LOGIC (Lógica de relatório integrada)
# =====================================================================================
def process_spreadsheet(input_path, holiday_day, texts):
    logs = []
    output_path = None
    report_data = {}
    excel_app = None
    
    try:
        excel_app = xw.App(visible=False)
        workbook = excel_app.books.open(input_path)
        sheet_name = '01. Calendario SCL Abarrotes'
        sheet = workbook.sheets[sheet_name]
        logs.append(texts['log_sheet_loaded'].format(file_name=os.path.basename(input_path)))

        delivery_columns = ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
        observations_column, store_col, carrier_col, chain_col = 'CT', 'F', 'B', 'D'
        weekday_map = {'L': 1, 'M': 2, 'W': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 7}

        holiday_col_letter = None
        for col_letter in delivery_columns:
            if sheet.range(f'{col_letter}3').value == holiday_day:
                holiday_col_letter = col_letter
                break
        
        if not holiday_col_letter:
            logs.append(texts['log_error_day_not_found'].format(holiday_day=holiday_day))
            return None, logs, None
        
        logs.append(texts['log_holiday_identified'].format(col_letter=holiday_col_letter))
        
        holiday_col_index = openpyxl.utils.column_index_from_string(holiday_col_letter)
        if holiday_col_index == openpyxl.utils.column_index_from_string(delivery_columns[0]):
            logs.append(texts['log_warning_first_day'])
            return None, logs, None

        previous_col_letter = openpyxl.utils.get_column_letter(holiday_col_index - 1)
        
        tasks_moved, rescheduled_stores, rescheduled_carriers, rescheduled_chains = 0, [], [], []
        max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        for row_index in range(8, max_row + 1):
            task_cell_range = sheet.range(f'{holiday_col_letter}{row_index}')
            if isinstance(task_cell_range.value, (int, float)) and 1 <= task_cell_range.value <= 6:
                new_weekday_number = weekday_map.get(str(sheet.range(f'{previous_col_letter}6').value).upper())
                if new_weekday_number:
                    # Executa o reagendamento
                    sheet.range(f'{previous_col_letter}{row_index}').value = new_weekday_number
                    task_cell_range.clear_contents()
                    sheet.range(f'{observations_column}{row_index}').value = texts['log_rescheduled_with_substitution'].format(holiday_day=holiday_day, col_letter=previous_col_letter)
                    
                    # Coleta dados para o relatório APENAS das linhas modificadas
                    tasks_moved += 1
                    rescheduled_stores.append(sheet.range(f'{store_col}{row_index}').value)
                    rescheduled_carriers.append(sheet.range(f'{carrier_col}{row_index}').value)
                    rescheduled_chains.append(sheet.range(f'{chain_col}{row_index}').value)
        
        logs.append(texts['log_rescheduling_complete'].format(tasks_moved=tasks_moved))
        
        # Monta o dicionário de relatório com os dados coletados
        report_data['rescheduled_deliveries'] = tasks_moved
        report_data['rescheduled_stores'] = sorted(list(set(s for s in rescheduled_stores if s)))

        by_carrier = dict(Counter(c for c in rescheduled_carriers if c).most_common(10))
        report_data['by_carrier'] = by_carrier
        report_data['max_carrier_count'] = max(by_carrier.values()) if by_carrier else 1
        report_data['total_carriers'] = len(set(c for c in rescheduled_carriers if c))

        by_chain = dict(Counter(c for c in rescheduled_chains if c).most_common(10))
        report_data['by_chain'] = by_chain
        report_data['max_chain_count'] = max(by_chain.values()) if by_chain else 1
        report_data['total_chains'] = len(set(c for c in rescheduled_chains if c))

        output_filename = f"{uuid.uuid4()}.xlsx"
        output_path = os.path.join(TMP_FOLDER, output_filename)
        workbook.save(output_path)
        return output_path, logs, report_data
    except Exception as e:
        logs.append(texts['log_error_read_sheet'].format(error=e))
        return None, logs, None
    finally:
        if excel_app: excel_app.quit()

# =====================================================================================
#  FLASK ROUTES (sem alterações, omitido por brevidade)
# =====================================================================================
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        session.pop('download_info', None)
        session.pop('report_data', None)

    lang = session.get('lang', 'en')
    if 'lang' in request.args:
        lang = request.args.get('lang')
        if lang in translations: session['lang'] = lang
        return redirect(url_for('index'))

    texts = translations[lang]

    if request.method == 'POST':
        file = request.files.get('spreadsheet')
        if not file or not file.filename:
            return render_template('index.html', texts=texts, error=texts['error_no_file_selected'])
        
        if file.filename.lower().endswith('.xlsx'):
            input_filename = f"{uuid.uuid4()}_{file.filename}"
            input_path = os.path.join(TMP_FOLDER, input_filename)
            file.save(input_path)

            holiday_day = int(request.form['holiday_day'])
            output_path, logs, report_data = process_spreadsheet(input_path, holiday_day, texts)
            os.remove(input_path)
            
            session['report_data'] = report_data or {}
            session['logs'] = logs or []
            
            if output_path:
                base_name, extension = os.path.splitext(file.filename)
                new_filename = f"{base_name}{texts['download_file_suffix']}{extension}"
                session['download_info'] = {'temp_filename': os.path.basename(output_path), 'final_filename': new_filename}
            
            return redirect(url_for('report_page'))
        else:
            return render_template('index.html', texts=texts, error=texts['error_wrong_file_type'])

    return render_template('index.html', texts=texts)

@app.route('/report')
def report_page():
    lang = session.get('lang', 'en')
    texts = translations[lang]
    report_data = session.get('report_data')
    logs = session.get('logs')
    download_info = session.get('download_info')

    if not report_data:
        return redirect(url_for('index'))

    return render_template('report.html', texts=texts, report=report_data, logs=logs, download_available=bool(download_info))

@app.route('/download')
def download_file():
    download_info = session.get('download_info')
    if download_info:
        temp_filename, final_filename = download_info['temp_filename'], download_info['final_filename']
        response = send_from_directory(directory=TMP_FOLDER, path=temp_filename, as_attachment=True, download_name=final_filename)
        @response.call_on_close
        def cleanup():
            try: os.remove(os.path.join(TMP_FOLDER, temp_filename))
            except Exception as e: print(f"Error removing file: {e}")
        return response
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)