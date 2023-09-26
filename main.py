from flask import Flask, render_template, request, redirect, url_for, flash, Response, send_file
import os, platform, subprocess
import openpyxl
from openpyxl import load_workbook
import requests
from fpdf import FPDF
from ics import Calendar, Event
import datetime
import sys
import re
from openpyxl import load_workbook
import os
import uuid
from flask import jsonify
import locale

app = Flask(__name__)
app.secret_key = "secret_key"  # Für Flash-Messages

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')

class ExcelProcessor:

    # Liste der Feiertage als Klassenvariable
    feiertage = [
        datetime.datetime(2023, 1, 1), datetime.datetime(2023, 4, 7), datetime.datetime(2023, 4, 10),
        datetime.datetime(2023, 5, 1), datetime.datetime(2023, 5, 18), datetime.datetime(2023, 5, 29),
        datetime.datetime(2023, 10, 3), datetime.datetime(2023, 10, 31), datetime.datetime(2023, 12, 25),
        datetime.datetime(2023, 12, 26), datetime.datetime(2024, 1, 1), datetime.datetime(2024, 3, 29),
        datetime.datetime(2024, 4, 1), datetime.datetime(2024, 5, 1), datetime.datetime(2024, 5, 9),
        datetime.datetime(2024, 5, 20), datetime.datetime(2024, 10, 3), datetime.datetime(2024, 10, 31),
        datetime.datetime(2024, 12, 25), datetime.datetime(2024, 12, 26), datetime.datetime(2025, 1, 1),
        datetime.datetime(2025, 4, 18), datetime.datetime(2025, 4, 21), datetime.datetime(2025, 5, 1),
        datetime.datetime(2025, 5, 29), datetime.datetime(2025, 6, 9), datetime.datetime(2025, 10, 3),
        datetime.datetime(2025, 10, 31), datetime.datetime(2025, 12, 25), datetime.datetime(2025, 12, 26)
    ]

    def __init__(self, filepath):
        self.filename = filepath  # Änderung hier
        self.workbook = load_workbook(filepath)

        # Define relevant sheets
        self.relevant_sheets = [sheet for sheet in self.workbook.sheetnames if
                                re.match(r'^KW\d{1,2}$', sheet)]  # Hinzugefügt

        # Create tempfiles directory if it doesn't exist
        if not os.path.exists('tempfiles'):
            os.makedirs('tempfiles')

        self.unique_id = uuid.uuid4().hex
        self.temp_pdf = f"tempfiles/temp_schedule_{self.unique_id}.pdf"
        self.modified_file = f"tempfiles/modified_file_{self.unique_id}.xlsx"

        # Aufrufen der cleanup Funktion bei der Initialisierung
        self.cleanup_old_files()

    def cleanup_old_files(self):
        # Liste der Ordner, in denen Dateien gelöscht werden sollen
        folders = ["tempfiles", "uploads"]

        # Aktuelles Datum und Zeitpunkt von vor 24 Stunden bestimmen
        now = datetime.datetime.now()
        one_day_ago = now - datetime.timedelta(days=1)

        for folder in folders:
            # Überprüfen, ob der Ordner existiert
            if os.path.exists(folder):
                print(f"Ordner {folder} wurde gefunden.")
                for filename in os.listdir(folder):
                    filepath = os.path.join(folder, filename)
                    # Nur Dateien überprüfen (keine Unterverzeichnisse)
                    if os.path.isfile(filepath):
                        file_creation_time = datetime.datetime.fromtimestamp(os.path.getctime(filepath))
                        print(f"Datei gefunden: {filepath}. Erstellt am: {file_creation_time}")

                        if file_creation_time < one_day_ago:
                            print(f"Die Datei {filepath} ist älter als 24 Stunden und wird gelöscht.")
                            os.remove(filepath)
                        else:
                            print(f"Die Datei {filepath} ist nicht älter als 24 Stunden.")
            else:
                print(f"Ordner {folder} wurde nicht gefunden!")

    def start_analysis(self):
        employees = set()
        if self.filename:
            self.workbook = openpyxl.load_workbook(self.modified_file)
            for sheet_name in self.relevant_sheets:
                sheet = self.workbook[sheet_name]
                letter_combinations = [sheet.cell(row=row, column=1).value for row in range(75, 136) if
                                       sheet.cell(row=row, column=1).value]
                for row in sheet.iter_rows(min_row=3, max_row=66, min_col=4, max_col=10):
                    for cell in row:
                        if cell.value:
                            for value in cell.value.split('/'):
                                for letter_combination in letter_combinations:
                                    if letter_combination[1:-1] in value:
                                        employees.add(letter_combination[1:-1])
        return sorted(employees)

    def replace_references_with_values(self):
        all_sheets = self.workbook.sheetnames
        for sheet_name in all_sheets:
            sheet = self.workbook[sheet_name]
            for row in sheet.iter_rows(min_row=3, max_row=66, min_col=4, max_col=10):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and re.match(r'^=.*$', cell.value):
                        reference = cell.value[1:]
                        if '!' in reference:
                            reference_sheet_name, reference_cell = reference.split('!')
                            reference_sheet_name = reference_sheet_name.strip("'")
                        else:
                            reference_sheet_name = sheet_name
                            reference_cell = reference
                        try:
                            reference_sheet = self.workbook[reference_sheet_name]
                            reference_value = reference_sheet[reference_cell].value
                            cell.value = reference_value
                        except KeyError:
                            print(f"Worksheet '{reference_sheet_name}' not found. Available sheets: {all_sheets}")
        self.workbook.save(self.modified_file)

    def check_plausibility(self):
        # Define the expected contents for each cell
        expected_contents = {
            'A15': 'Eingriff',
            'A24': 'POB',
            'A30': 'Inten',
            'A40': 'BD 1',
            'A41': 'BD 2',
        }

        # Iterate over the relevant sheets
        for sheet_name in self.relevant_sheets:
            # Get the sheet from the workbook
            sheet = self.workbook[sheet_name]

            # Iterate over the expected contents
            for cell, expected_content in expected_contents.items():
                # Get the actual content of the cell
                actual_content = sheet[cell].value

                # If the actual content does not contain the expected content, show an error message and return False
                if not actual_content or expected_content not in actual_content:
                    print("Fehler: Offenbar wurde die Sortierung der Wochenpläne verändert oder Sie haben einen alten Dienstlan geladen. Eine verlässliche Extraktion der Dienste ist nicht gewährleistet. Bitte informieren Sie den Entwickler")
                    return False

        # If all the checks passed, return True
        return True

    def show_schedule(self, selected_employee):

        # Setze die Lokalisierung auf Deutsch
        locale.setlocale(locale.LC_TIME, 'de_DE.UTF-8')

        # Get the covered days from the workbook
        covered_days = get_covered_days(self.workbook, self.relevant_sheets)

        # Get the month and year from the tenth covered day
        tenth_day = datetime.datetime.strptime(covered_days[9].split()[0], '%d.%m.%Y')
        month_year = tenth_day.strftime('%B %Y')

        # Initialize a dictionary to store the schedule of the selected employee
        employee_schedule = {}

        # Iterate over the relevant sheets
        for sheet_name in self.relevant_sheets:
            # Get the sheet from the workbook
            sheet = self.workbook[sheet_name]

            # Iterate over the columns of the sheet
            for col in range(4, 11):
                # Get the date from the second row of the sheet
                date = sheet.cell(row=2, column=col).value

                # If the date is covered, get the service of the selected employee and add it to the employee_schedule dictionary
                if date and date.strftime('%d.%m.%Y %A') in covered_days:
                    service = self.get_service(sheet, selected_employee, col)
                    if date in self.feiertage:
                        service += ' (Feiertag)'
                    employee_schedule[date.strftime('%d.%m.%Y')] = service

        # Initialize a string to store the schedule of the selected employee
        schedule_text = f"Dienstplan für {selected_employee} - {month_year}\n\n"

        # Define a list of weekdays in German
        weekdays_german = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag']

        # Append the formatted employee schedule to the schedule_text string
        for i, (date, service) in enumerate(employee_schedule.items()):
            # Get the weekday in German
            weekday = weekdays_german[datetime.datetime.strptime(date, '%d.%m.%Y').weekday()]
            line = f"{date} ({weekday}): {service}\n"
            schedule_text += line

            # Insert a line after every Sunday
            if weekday == 'Sonntag':
                schedule_text += '-' * 50 + '\n'

        # Insert the disclaimer text
        disclaimer = "\nBitte überprüft den Plan auf seine Richtigkeit. Achtet im offiziellen Plan auf kurzfristige Änderungen!\n"
        schedule_text += disclaimer

        return schedule_text

    def get_service(self, sheet, employee, col):
        # Define a dictionary to map the row numbers to the services
        services = {
            3: 'OP-Koordination',
            4: 'FD-OP',
            5: 'FD-OP',
            6: 'FD-OP',
            7: 'FD-OP',
            8: 'FD-OP',
            9: 'FD-OP',
            10: 'FD-OP',
            11: 'FD-OP',
            12: 'FD-lang',
            14: 'FD-Außenbezirke',
            15: 'EGR-OA',
            16: 'FD-EGR',
            17: 'FD-EGR',
            18: 'FD-EGR',
            19: 'FD-BronchoHKL',
            20: 'FD-Geb',
            # The service for row 21 will be set later based on the date condition
            22: 'SD11',
            23: 'SD13',
            24: 'POBE',
            25: 'Prämed',
            26: 'Prämed',
            27: 'Prämed',
            28: 'OA-ZOP',
            29: 'FD-Einarbeitung',
            30: 'FD-OA-Int',
            31: 'FD-FA-Int',
            32: 'FD-Int',
            33: 'FD-Int',
            34: 'SD-Int',
            35: 'SD-Int',
            36: 'ND-Int',
            37: 'NEF-Tag',
            38: 'NEF-Nacht',
            39: 'ITW',
            40: 'BD1',
            41: 'BD2',
            42: 'Rufdienst',
        }

        # Add 'Frei' service for rows 43 to 73
        for row in range(43, 74):
            services[row] = 'Frei'

        # Get the date from the current column and determine the service for row 12
        current_date = sheet.cell(row=2, column=col).value

        # Check if the current date is a Friday
        if current_date and current_date.weekday() == 4:  # Wenn es Freitag ist
            services[12] = 'FD10'
        else:
            services[12] = 'FD-lang'

        if current_date and current_date.month <= 9:
            services[21] = 'SD11'
        else:
            services[21] = 'SD930'

        # Initialize a list to store the services of the employee
        employee_services = []

        # Iterate over the rows of the sheet
        for row in range(3, 74):
            # Get the cell value
            cell_value = sheet.cell(row=row, column=col).value

            # If the cell value contains the employee name, append the service to the employee_services list
            if cell_value:
                for value in cell_value.split('/'):
                    if employee in value:
                        service = services[row]
                        if row in [40, 41] and '/' in cell_value:
                            if cell_value.index(employee) == 0:
                                service += '/Tag'
                            else:
                                service += '/Nacht'
                        employee_services.append(service)

        # Wenn der Mitarbeitername in keiner Zeile gefunden wird
        if not employee_services:
            date = sheet.cell(row=2, column=col).value

            # Überprüfen, ob das Datum ein Samstag, Sonntag oder ein Feiertag ist
            if date and (date.weekday() in [5, 6] or date in self.feiertage):
                employee_services.append('Frei')
            else:
                employee_services.append('?')

        return ', '.join(employee_services)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Keine Datei ausgewählt')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('Keine Datei ausgewählt')
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)

        # Load the workbook
        workbook = load_workbook(filename)

        # Define the relevant sheets based on your provided pattern
        relevant_sheets = [sheet for sheet in workbook.sheetnames if re.match(r'^KW\d{1,2}$', sheet)]

        # Create an instance of ExcelProcessor and replace the references
        processor = ExcelProcessor(filename)
        processor.relevant_sheets = relevant_sheets
        processor.replace_references_with_values()

        # Start the analysis to get employees
        employees = processor.start_analysis()

        # Anstatt die HTML-Seite direkt zurückzugeben, geben Sie die UID und die Mitarbeiterliste als JSON zurück:
        return jsonify(uid=processor.unique_id, employees=employees)

    else:
        flash('Nicht unterstütztes Dateiformat')
        return redirect(request.url)



def get_covered_days(workbook, relevant_sheets):
    # Initialize a list to store the covered days
    covered_days = []

    # Iterate over the relevant sheets
    for sheet_name in relevant_sheets:
        # Get the sheet from the workbook
        sheet = workbook[sheet_name]

        # Get the dates from the second row of the sheet
        dates = [sheet.cell(row=2, column=col).value for col in range(4, 11)]

        # Append the dates to the covered_days list
        covered_days.extend(dates)

        # Convert the dates to the desired format
    covered_days = [date.strftime('%d.%m.%Y %A') for date in covered_days if isinstance(date, datetime.datetime)]

    return covered_days


@app.route('/handle_employee_selection', methods=['POST'])
def handle_employee_selection():
    data = request.json
    selected_employee = data.get('employee')
    uid = data.get('uid')

    # Drucken Sie den ausgewählten Mitarbeiter in die Konsole
    print(f"Ausgewählter Mitarbeiter: {selected_employee}")

    # Erstellen Sie den Prozessor mit der Datei, die der UID entspricht:
    filepath = f"tempfiles/modified_file_{uid}.xlsx"
    processor = ExcelProcessor(filepath)

    # Überprüfen Sie die Plausibilität
    if not processor.check_plausibility():
        error_message = ("Fehler: Offenbar wurde die Sortierung der Wochenpläne verändert oder "
                         "Sie haben einen alten Dienstlan geladen. Eine verlässliche Extraktion der Dienste ist "
                         "nicht gewährleistet. Bitte informieren Sie den Entwickler")
        return jsonify(message="Fehler bei der Plausibilitätsprüfung", schedule=error_message)

    # Get the schedule for the selected employee
    schedule = processor.show_schedule(selected_employee)

    # Return the schedule in the response
    return jsonify(message="Erfolgreich verarbeitet", schedule=schedule)


@app.route('/generate_pdf', methods=['POST'])
def generate_pdf():
    # Sie können hier Daten aus dem Post-Body extrahieren, falls benötigt.
    # z.B.: data = request.json
    # Der eigentliche Inhalt, der ins PDF soll (z.B. Dienstplan), sollte wahrscheinlich vom Frontend übertragen werden.

    # Erstellen Sie das PDF (ich übernehme den größten Teil Ihres Codes)
    schedule_text = request.json.get("schedule_text", "")

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=8)

    lines = schedule_text.split('\n')
    for line in lines:
        if 'Samstag' in line or 'Sonntag' in line:
            pdf.set_fill_color(200, 200, 200)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.multi_cell(0, 5, line, fill=True)

    creation_date = datetime.datetime.now().strftime('%d.%m.%Y')   # Hier wurde die Änderung vorgenommen
    creation_time = datetime.datetime.now().strftime('%H:%M')      # Hier wurde die Änderung vorgenommen
    pdf.multi_cell(0, 10, f"Erstellt am {creation_date} um {creation_time} Uhr", align='L')

    # Statt es auf die Festplatte zu speichern, speichern wir es in den Speicher.
    pdf_content = pdf.output(dest='S').encode('latin1')

    firstLine = schedule_text.splitlines()[0]
    log_action(firstLine, "PDF erstellt")

    # Antwort mit PDF-Inhalt
    response = Response(pdf_content, content_type='application/pdf')
    return response


@app.route('/generate_ics', methods=['POST'])
def generate_ics():
    # Die ics_export_action-Funktion hier umformatieren:

    data = request.json
    schedule_text = data['schedule_text']

    c = Calendar()

    lines = schedule_text.split('\n')
    for line in lines:
        parts = line.split(": ")
        if len(parts) == 2:
            date_str, service = parts
            date_parts = date_str.split(" ")
            if len(date_parts) == 2:
                date, weekday = date_parts
                day, month, year = date.split('.')
                formatted_date = f"{year}-{month}-{day}"
                e = Event()
                e.name = service
                e.begin = formatted_date
                e.make_all_day()
                c.events.add(e)

    employee_name = lines[0].split(' ')[2] if len(lines) > 0 else "Unbekannt"
    month_year = ' '.join(lines[0].split(' ')[-2:]) if len(lines) > 0 else "Unbekannt"
    ics_filename = f"Dienstplan-{employee_name}-{month_year}.ics"

    with open(ics_filename, 'w', encoding='utf-8') as my_file:
        my_file.writelines(str(c))

    firstLine = schedule_text.splitlines()[0]
    log_action(firstLine, "ICS erstellt")

    return send_file(ics_filename, as_attachment=True)


def log_action(schedule_first_line, filetype):
    # IFTTT-Webhook-URL
    webhook_url = 'https://maker.ifttt.com/trigger/Dienstplan/with/key/jeYCdecKNOCcFc5NRUWu07ufuHSQKwrnyCuA_eQZZN7'

    # Daten für den Webhook
    data = {
        "value1": schedule_first_line,
        "value2": filetype
    }

    # Senden des Webhooks
    response = requests.post(webhook_url, json=data) # Verwenden Sie json=data, um es als JSON zu senden
    return response



if __name__ == '__main__':
    app.run(debug=True)



if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
