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

translations = {
    'en': {
        'title': "🤖 Automatic Holiday Rescheduler",
        'description': "This tool automates the rescheduling of deliveries. Upload your spreadsheet, enter the holiday day, and get the modified file.",
        'file_uploader_label': "1. Choose your scheduling spreadsheet",
        'number_input_label': "2. Enter the holiday day",
        'button_label': "Reschedule Spreadsheet",
        'processing_text': "Processing...",
        'report_header': "Operation Report",
        'success_message': "Your spreadsheet has been successfully rescheduled!",
        'download_button_label': "Download Rescheduled Spreadsheet",
        'view_report_button_label': "View Spreadsheet Report",
        'report_title': "Spreadsheet Analysis Report",
        'total_deliveries': "Total Deliveries in Period",
        'deliveries_by_carrier': "Deliveries by Carrier",
        'deliveries_by_chain': "Deliveries by Chain",
        'close_button': "Close",
        'download_file_suffix': "_rescheduled",
        'error_upload_file': "Please upload a spreadsheet to continue.",
        'error_no_file_part': "No file part in the request.",
        'error_no_file_selected': "No file selected.",
        'log_sheet_loaded': "Spreadsheet '{file_name}' loaded successfully.",
        'log_error_read_sheet': "ERROR: Could not read the spreadsheet. Please check if it is the correct file. Details: {error}",
        'log_error_day_not_found': "ERROR: The day {holiday_day} was not found in row 3 of the Delivery columns.",
        'log_holiday_identified': "Holiday identified in the Delivery column: {col_letter}",
        'log_warning_first_day': "Warning: The holiday is the first day of the period. It cannot be anticipated.",
        'log_rescheduled_with_substitution': "Delivery rescheduled (with substitution) from day {holiday_day} to column {col_letter}.",
        'log_rescheduling_complete': "Rescheduling completed. {tasks_moved} tasks were moved.",
    },
    'es': {
        'title': "🤖 Reprogramador Automático de Feriados",
        'description': "Esta herramienta automatiza la reprogramación de entregas. Suba su planilla, ingrese el día feriado y obtenga el archivo modificado.",
        'file_uploader_label': "1. Elija su planilla de programación",
        'number_input_label': "2. Ingrese el día feriado",
        'button_label': "Reprogramar Planilla",
        'processing_text': "Procesando...",
        'report_header': "Reporte de Operación",
        'success_message': "¡Su planilla ha sido reprogramada exitosamente!",
        'download_button_label': "Descargar Planilla Reprogramada",
        'view_report_button_label': "Ver Reporte de la Planilla",
        'report_title': "Reporte de Análisis de la Planilla",
        'total_deliveries': "Total de Entregas en el Período",
        'deliveries_by_carrier': "Entregas por Transportista",
        'deliveries_by_chain': "Entregas por Cadena",
        'close_button': "Cerrar",
        'download_file_suffix': "_reprogramada",
        'error_upload_file': "Por favor, suba una planilla para continuar.",
        'error_no_file_part': "No hay archivo en la solicitud.",
        'error_no_file_selected': "No se ha seleccionado ningún archivo.",
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
#  FUNÇÃO DE RELATÓRIO ATUALIZADA
# =====================================================================================
def generate_report(sheet):
    report_data = {}
    
    carrier_col = 'B'
    chain_col = 'D'
    delivery_cols = ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
    
    max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    
    carriers = []
    chains = []
    total_deliveries = 0
    
    # Usamos .options(pd.DataFrame) para ler os dados de forma otimizada
    # Lendo um range maior para garantir que todos os dados sejam capturados
    data_range = sheet.range(f'A8:AN{max_row}').options(ndim=2).value
    
    for row_data in data_range:
        carrier = row_data[1]  # Coluna B é o índice 1
        if carrier:
            carriers.append(carrier)
            
        chain = row_data[3]  # Coluna D é o índice 3
        if chain:
            chains.append(chain)
        
        # Colunas AI a AN (índices 34 a 39)
        for i in range(34, 40):
            delivery_value = row_data[i]
            if isinstance(delivery_value, (int, float)) and delivery_value > 0:
                total_deliveries += 1

    report_data['total_deliveries'] = total_deliveries
    
    by_carrier = dict(Counter(carriers).most_common(10))
    report_data['by_carrier'] = by_carrier
    # Adiciona o valor máximo para cálculo da barra de progresso no frontend
    report_data['max_carrier_count'] = max(by_carrier.values()) if by_carrier else 1

    by_chain = dict(Counter(chains).most_common(10))
    report_data['by_chain'] = by_chain
    # Adiciona o valor máximo para cálculo da barra de progresso no frontend
    report_data['max_chain_count'] = max(by_chain.values()) if by_chain else 1
    
    return report_data

# =====================================================================================
#  BUSINESS LOGIC e FLASK ROUTES (sem alterações, omitido por brevidade)
# =====================================================================================
def process_spreadsheet(input_path, holiday_day, texts):
    logs = []
    output_path = None
    report_data = None
    excel_app = None
    
    try:
        excel_app = xw.App(visible=False)
        workbook = excel_app.books.open(input_path)
        sheet_name = '01. Calendario SCL Abarrotes'
        sheet = workbook.sheets[sheet_name]
        logs.append(texts['log_sheet_loaded'].format(file_name=os.path.basename(input_path)))

        report_data = generate_report(sheet)

        delivery_columns = ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
        observations_column = 'CT'
        weekday_map = {'L': 1, 'M': 2, 'W': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 7}

        holiday_col_letter = None
        for col_letter in delivery_columns:
            day_in_sheet = sheet.range(f'{col_letter}3').value
            if day_in_sheet == holiday_day:
                holiday_col_letter = col_letter
                break
        
        if not holiday_col_letter:
            logs.append(texts['log_error_day_not_found'].format(holiday_day=holiday_day))
            return None, logs, None

        logs.append(texts['log_holiday_identified'].format(col_letter=holiday_col_letter))
        
        holiday_col_index = openpyxl.utils.column_index_from_string(holiday_col_letter)
        if holiday_col_index == openpyxl.utils.column_index_from_string(delivery_columns[0]):
            logs.append(texts['log_warning_first_day'])
            return None, logs, report_data

        previous_col_index = holiday_col_index - 1
        previous_col_letter = openpyxl.utils.get_column_letter(previous_col_index)
        
        tasks_moved = 0
        max_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

        for row_index in range(8, max_row + 1):
            task_cell_range = sheet.range(f'{holiday_col_letter}{row_index}')
            task_value = task_cell_range.value
            if isinstance(task_value, (int, float)) and 1 <= task_value <= 6:
                weekday_initial = str(sheet.range(f'{previous_col_letter}6').value).upper()
                new_weekday_number = weekday_map.get(weekday_initial)
                
                if new_weekday_number:
                    sheet.range(f'{previous_col_letter}{row_index}').value = new_weekday_number
                    task_cell_range.clear_contents()
                    log_message = texts['log_rescheduled_with_substitution'].format(holiday_day=holiday_day, col_letter=previous_col_letter)
                    sheet.range(f'{observations_column}{row_index}').value = log_message
                    tasks_moved += 1
        
        logs.append(texts['log_rescheduling_complete'].format(tasks_moved=tasks_moved))

        output_filename = f"{uuid.uuid4()}.xlsx"
        output_path = os.path.join(TMP_FOLDER, output_filename)
        workbook.save(output_path)
        return output_path, logs, report_data

    except Exception as e:
        logs.append(texts['log_error_read_sheet'].format(error=e))
        return None, logs, None
    finally:
        if excel_app:
            excel_app.quit()


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        session.pop('download_info', None)
        session.pop('report_data', None)

    lang = session.get('lang', 'en')
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
            input_filename = f"{uuid.uuid4()}_{file.filename}"
            input_path = os.path.join(TMP_FOLDER, input_filename)
            file.save(input_path)

            holiday_day = int(request.form['holiday_day'])
            output_path, logs, report_data = process_spreadsheet(input_path, holiday_day, texts)
            
            os.remove(input_path)
            
            if report_data:
                session['report_data'] = report_data
            
            if output_path:
                base_name, extension = os.path.splitext(file.filename)
                new_filename = f"{base_name}{texts['download_file_suffix']}{extension}"

                session['download_info'] = {
                    'temp_filename': os.path.basename(output_path),
                    'final_filename': new_filename
                }
                
                return render_template('index.html', texts=texts, logs=logs, success=True)
            else:
                return render_template('index.html', texts=texts, logs=logs, success=False)

    return render_template('index.html', texts=texts)


@app.route('/download')
def download_file():
    download_info = session.get('download_info', None)
    if download_info:
        temp_filename = download_info['temp_filename']
        final_filename = download_info['final_filename']
        
        response = send_from_directory(
            directory=TMP_FOLDER,
            path=temp_filename,
            as_attachment=True,
            download_name=final_filename
        )
        
        @response.call_on_close
        def cleanup():
            try:
                os.remove(os.path.join(TMP_FOLDER, temp_filename))
            except Exception as e:
                print(f"Error removing file: {e}")

        return response
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)