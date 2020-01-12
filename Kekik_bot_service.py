from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime


def check_disk():
    """
    Смотрим файлы на диске
    :return: Объект drive
    """
    return service_for_drive.files().list(pageSize=10,
                                          fields="nextPageToken, files(id, name, mimeType, parents, createdTime, "
                                                 "permissions, quotaBytesUsed)").execute()


def create_spreadsheet():
    """
    Создаем таблицу
    :return: Объект spreadsheet
    """
    months = ('Январь', 'Февраль',
              'Март', 'Апрель', 'Май',
              'Июнь', 'Июль', 'Август',
              'Сентябрь', 'Октябрь', 'Ноябрь',
              'Декабрь',)
    sheets = [{'properties': {'sheetType': 'GRID',
                              'sheetId': 0,
                              'title': 'Сводная таблица по расходам и доходам по категориям',
                              'gridProperties': {'rowCount': 0, 'columnCount': 50}}}, ]
    for month_number, month in enumerate(months):
        sheets.append({'properties': {'sheetType': 'GRID',
                                      'sheetId': month_number + 1,
                                      'title': month,
                                      'gridProperties': {'rowCount': 0, 'columnCount': 0}}})
    creating_spreadsheet_body = {
        'properties': {'title': 'Бюджет с Кекиком', 'locale': 'ru_RU'},
        'sheets': sheets
    }
    return service_for_sheets.spreadsheets().create(body=creating_spreadsheet_body).execute()


def permissions_for_owner(user_gmail):
    """
    "Дарим" созданную программой таблицу пользователю: меняем его статус на владельца
    :return: None
    """
    gmail = user_gmail  # Пользователь вводит свой адрес
    request = check_disk()
    for files in request['files']:
        if files['name'] == 'Бюджет с Кекиком' and files['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            spreadsheet_id = files['id']  # ID таблицы, с которой будем работать
            shareRes = service_for_drive.permissions().create(
                fileId=spreadsheet_id,
                body={'type': 'user', 'role': 'owner', 'emailAddress': gmail},
                transferOwnership=True
            ).execute()
    print('Таблица "Бюджет с Кекиком" успешна создана, теперь вы являетесь ее владельцем.'
          'Пожалуйся, не закрывайте мне доступ к ней!')


def permissions_for_second_user(user_gmail):
    gmail = user_gmail
    request = check_disk()
    for files in request['files']:
        if files['name'] == 'Бюджет с Кекиком' and files['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            spreadsheet_id = files['id']  # ID таблицы, с которой будем работать
            shareRes = service_for_drive.permissions().create(
                fileId=spreadsheet_id,
                body={'type': 'user', 'role': 'writer', 'emailAddress': gmail},
                transferOwnership=False
            ).execute()
    print('Теперь вы можете редактировать таблицу "Бюджет с Кекиком".')


def get_spreadsheet_id():
    """
    Получение ID нужной таблицы
    :return: ID таблицы с именем 'Бюджет с Кекиком'
    """
    request = check_disk()
    for files in request['files']:
        if files['name'] == 'Бюджет с Кекиком' and files['mimeType'] == 'application/vnd.google-apps.spreadsheet':
            return files['id']


class Spreadsheet:
    """
    Класс обертка для удобного доступа к API
    Источник https://habr.com/ru/post/305378/
    """

    def __init__(self):
        self.valueRanges = []
        self.sheetTitles = []
        self.requests = []

    def request_from_spreadsheet(self, valueRange=None, includeGridData=False):
        self.valueRange = valueRange
        self.includeGridData = includeGridData
        return service_for_sheets.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=self.valueRange,
                                                     includeGridData=self.includeGridData).execute()

    def sheetsTitles_list(self):
        request = self.request_from_spreadsheet()
        for sheets in request['sheets']:
            self.sheetTitles.append(sheets['properties']['title'])

    def prepare_setValues(self, cellsRange, values, sheet_number, majorDimension="ROWS"):
        self.valueRanges.append(
            {"range": self.sheetTitles[sheet_number] + "!" + cellsRange, "majorDimension": majorDimension,
             "values": values})

    def runPrepared(self, valueInputOption="USER_ENTERED"):
        if len(self.valueRanges) > 0:
            upd1Res = service_for_sheets.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id,
                                                                             body={"valueInputOption": valueInputOption,
                                                                                   "data": self.valueRanges}).execute()
        if len(self.requests) > 0:
            upd2Res = service_for_sheets.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                                    body={"requests": self.requests}).execute()
        self.valueRanges = []
        self.requests = []

    def toGridRange(self, cellsRange, sheet_number):
        if isinstance(cellsRange, str):
            startCell, endCell = cellsRange.split(":")
            cellsRange = {}
            rangeAZ = range(ord('A'), ord('z') + 1)
            if startCell[1].isalpha():
                startCell = startCell.lower()
            if endCell[1].isalpha():
                endCell = endCell.lower()
            if ord(startCell[0]) in rangeAZ:
                if ord(startCell[0]) > 90:
                    cellsRange["startColumnIndex"] = ord(startCell[1]) - ord('A') - 6
                    startCell = startCell[2:]
                else:
                    cellsRange["startColumnIndex"] = ord(startCell[0]) - ord('A')
                    startCell = startCell[1:]
            if ord(endCell[0]) in rangeAZ:
                if ord(endCell[0]) > 90:
                    cellsRange["endColumnIndex"] = ord(endCell[1]) - ord('A') - 5
                    endCell = endCell[2:]
                else:
                    cellsRange["endColumnIndex"] = ord(endCell[0]) - ord('A') + 1
                    endCell = endCell[1:]
            if len(startCell) > 0:
                cellsRange["startRowIndex"] = int(startCell) - 1
            if len(endCell) > 0:
                cellsRange["endRowIndex"] = int(endCell)
        cellsRange["sheetId"] = sheet_number
        return cellsRange

    def prepare_mergeCells(self, cellsRange, sheet_number, mergeType="MERGE_ALL"):
        self.requests.append(
            {"mergeCells": {"range": self.toGridRange(cellsRange, sheet_number), "mergeType": mergeType}})

    # formatJSON should be dict with userEnteredFormat to be applied to each cell
    def prepare_setCellsFormat(self, cellsRange, formatJSON, sheet_number, fields="userEnteredFormat"):
        self.requests.append({"repeatCell": {"range": self.toGridRange(cellsRange, sheet_number),
                                             "cell": {"userEnteredFormat": formatJSON}, "fields": fields}})

    # formatsJSON should be list of lists of dicts with userEnteredFormat for each cell in each row
    def prepare_setCellsFormats(self, cellsRange, formatsJSON, sheet_number, fields="userEnteredFormat"):
        self.requests.append({"updateCells": {"range": self.toGridRange(cellsRange, sheet_number),
                                              "rows": [{"values": [{"userEnteredFormat": cellFormat} for cellFormat
                                                                   in rowFormats]} for rowFormats in formatsJSON],
                                              "fields": fields}})

    def prepare_setBorder_bot(self, cellsRange, sheet_number, width=1):
        self.requests.append({"updateBorders": {"range": self.toGridRange(cellsRange, sheet_number),
                                                "bottom": {
                                                    'style': 'SOLID',
                                                    'width': width,
                                                    'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}
                                                }}})

    def prepare_setBorder_top(self, cellsRange, sheet_number, width=1):
        self.requests.append({"updateBorders": {"range": self.toGridRange(cellsRange, sheet_number),
                                                "top": {
                                                    'style': 'SOLID',
                                                    'width': width,
                                                    'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}
                                                }}})

    def prepare_setBorder_left(self, cellsRange, sheet_number, width=1):
        self.requests.append({"updateBorders": {"range": self.toGridRange(cellsRange, sheet_number),
                                                "left": {
                                                    'style': 'SOLID',
                                                    'width': width,
                                                    'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}
                                                }}})

    def prepare_setBorder_right(self, cellsRange, sheet_number, width=1):
        self.requests.append({"updateBorders": {"range": self.toGridRange(cellsRange, sheet_number),
                                                "right": {
                                                    'style': 'SOLID',
                                                    'width': width,
                                                    'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}
                                                }}})

    def update_cell_data(self, data, date, user_name):
        value = datetime.datetime.fromtimestamp(date)
        day = value.strftime('%d')
        month = value.strftime('%m')
        self.sheetsTitles_list()
        category = ''  # TODO переделать костыль
        for _ in data.split():
            if _.isdigit():
                value = _
            else:
                category += f' {_}'  # TODO переделать костыль

        category = category.strip()  # TODO переделать костыль
        current_category = define_category(category)
        sheet = int(month)
        row = int(day) + 2
        user_1_columns = ('B', 'E', 'H', 'K', 'N', 'Q', 'T', 'W', 'Z', 'AC', 'AF', 'AI',)
        user_2_columns = ('C', 'F', 'I', 'L', 'O', 'R', 'U', 'X', 'AA', 'AD', 'AG', 'AJ',)
        category_row = current_category[2]

        if user_name == user_1_name:
            profit_column = 'B'
            loss_column = 'E'
            comment_column = 'K'
            for index, column in enumerate(user_1_columns):
                if index + 1 == sheet:
                    category_column = column
                    break

        elif user_name == user_2_name:
            profit_column = 'C'
            loss_column = 'F'
            comment_column = 'L'
            for index, column in enumerate(user_2_columns):
                if index + 1 == sheet:
                    category_column = column
                    break

        if current_category[0] == True:
            request = self.request_from_spreadsheet(
                valueRange=f'{self.sheetTitles[sheet]}!{profit_column}{row}:{profit_column}{row}',
                includeGridData=True)
            data_availability_check = request['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            if 'formattedValue' in data_availability_check:
                self.prepare_setValues(f'{profit_column}{row}:{profit_column}{row}',
                                       [[f'={data_availability_check["formattedValue"]}+{value}']], sheet)
            else:
                self.prepare_setValues(f'{profit_column}{row}:{profit_column}{row}', [[f'={value}']], sheet)

            request = self.request_from_spreadsheet(
                valueRange=f'{self.sheetTitles[sheet]}!{comment_column}{row}:{comment_column}{row}',
                includeGridData=True)
            data_availability_check = request['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            if 'formattedValue' in data_availability_check:
                self.prepare_setValues(f'{comment_column}{row}:{comment_column}{row}',
                                       [[f'{data_availability_check["formattedValue"]}, {category.lower()}']], sheet)
            else:
                self.prepare_setValues(f'{profit_column}{row}:{profit_column}{row}', [[f'{category.capitalize()}']],
                                       sheet)

        elif current_category[0] == False:
            request = self.request_from_spreadsheet(
                valueRange=f'{self.sheetTitles[sheet]}!{loss_column}{row}:{loss_column}{row}',
                includeGridData=True)
            data_availability_check = request['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            if 'formattedValue' in data_availability_check:
                self.prepare_setValues(f'{loss_column}{row}:{loss_column}{row}',
                                       [[f'={data_availability_check["formattedValue"]}+{value}']], sheet)
            else:
                self.prepare_setValues(f'{loss_column}{row}:{loss_column}{row}', [[f'={value}']], sheet)

            request = self.request_from_spreadsheet(
                valueRange=f'{self.sheetTitles[sheet]}!{comment_column}{row}:{comment_column}{row}',
                includeGridData=True)
            data_availability_check = request['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            if 'formattedValue' in data_availability_check:
                self.prepare_setValues(f'{comment_column}{row}:{comment_column}{row}',
                                       [[f'{data_availability_check["formattedValue"]}, {category.lower()}']], sheet)
            else:
                self.prepare_setValues(f'{comment_column}{row}:{comment_column}{row}', [[f'{category.capitalize()}']],
                                       sheet)

            request = self.request_from_spreadsheet(
                valueRange=f'{self.sheetTitles[0]}!{category_column}{category_row}:{category_column}{category_row}',
                includeGridData=True)
            data_availability_check = request['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            if 'formattedValue' in data_availability_check:
                self.prepare_setValues(f'{category_column}{category_row}:{category_column}{category_row}',
                                       [[f'={data_availability_check["formattedValue"]}+{value}']], 0)
            else:
                self.prepare_setValues(f'{category_column}{category_row}:{category_column}{category_row}',
                                       [[f'={value}']], 0)
        self.runPrepared()


def prepare_sheets(spreadsheet_redactor):
    spreadsheet_redactor.prepare_mergeCells('A1:A2', 0)
    spreadsheet_redactor.prepare_mergeCells('B1:D1', 0)
    spreadsheet_redactor.prepare_mergeCells('E1:G1', 0)
    spreadsheet_redactor.prepare_mergeCells('H1:J1', 0)
    spreadsheet_redactor.prepare_mergeCells('K1:M1', 0)
    spreadsheet_redactor.prepare_mergeCells('N1:P1', 0)
    spreadsheet_redactor.prepare_mergeCells('Q1:S1', 0)
    spreadsheet_redactor.prepare_mergeCells('T1:V1', 0)
    spreadsheet_redactor.prepare_mergeCells('W1:Y1', 0)
    spreadsheet_redactor.prepare_mergeCells('Z1:AB1', 0)
    spreadsheet_redactor.prepare_mergeCells('AC1:AE1', 0)
    spreadsheet_redactor.prepare_mergeCells('AF1:AH1', 0)
    spreadsheet_redactor.prepare_mergeCells('AI1:AK1', 0)
    spreadsheet_redactor.runPrepared()
    for sheet_id in range(12):
        spreadsheet_redactor.prepare_mergeCells("A1:A2", sheet_id + 1)
        spreadsheet_redactor.prepare_mergeCells("B1:D1", sheet_id + 1)
        spreadsheet_redactor.prepare_mergeCells('E1:G1', sheet_id + 1)
        spreadsheet_redactor.prepare_mergeCells('H1:H2', sheet_id + 1)
        spreadsheet_redactor.prepare_mergeCells('I1:J1', sheet_id + 1)
        spreadsheet_redactor.prepare_mergeCells('K1:L1', sheet_id + 1)
        spreadsheet_redactor.runPrepared()


def create_template(spreadsheet_redactor):
    user_1 = user_1_name
    user_2 = user_2_name
    spreadsheet_redactor.sheetsTitles_list()
    spreadsheet_redactor.prepare_setValues("A1:A1", [["Категории"]], 0)
    spreadsheet_redactor.prepare_setValues("A3:A24", [
        ["Продукты", "Красота", "Медицина", "Аптека", "Одежда", "Развлечения", "Спорт", "Связь", "Транспорт",
         "Услуги банка", "Домашние животные", "ЖКХ", "Сервисы(музыка/сериалы/игры)", "Кафе и рестораны", "Переводы",
         "Ипотека/Аренда", "Техника", "Ремонт", "Для дома", "Подарки", "Разное", "Всего"]], 0, "COLUMNS")
    spreadsheet_redactor.prepare_setValues("B2:AK2", [
        [user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', user_1,
         user_2, 'Сумма', user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', user_1, user_2,
         'Сумма', user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', user_1, user_2, 'Сумма']], 0)
    spreadsheet_redactor.prepare_setValues('B1:B1', [['Январь']], 0)
    spreadsheet_redactor.prepare_setValues('E1:E1', [['Февраль']], 0)
    spreadsheet_redactor.prepare_setValues('H1:H1', [['Март']], 0)
    spreadsheet_redactor.prepare_setValues('K1:K1', [['Апрель']], 0)
    spreadsheet_redactor.prepare_setValues('N1:N1', [['Май']], 0)
    spreadsheet_redactor.prepare_setValues('Q1:Q1', [['Июнь']], 0)
    spreadsheet_redactor.prepare_setValues('T1:T1', [['Июль']], 0)
    spreadsheet_redactor.prepare_setValues('W1:W1', [['Август']], 0)
    spreadsheet_redactor.prepare_setValues('Z1:Z1', [['Сентябрь']], 0)
    spreadsheet_redactor.prepare_setValues('AC1:AC1', [['Октябрь']], 0)
    spreadsheet_redactor.prepare_setValues('AF1:AF1', [['Ноябрь']], 0)
    spreadsheet_redactor.prepare_setValues('AI1:AI1', [['Декабрь']], 0)
    spreadsheet_redactor.prepare_setCellsFormat('A1:A24',
                                                {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE',
                                                 'textFormat': {'bold': True}}, 0)
    spreadsheet_redactor.prepare_setCellsFormat('B1:AK2',
                                                {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE',
                                                 'textFormat': {'bold': True}}, 0)
    spreadsheet_redactor.prepare_setBorder_bot('A1:A2', 0, 3)
    spreadsheet_redactor.prepare_setBorder_bot('A23:AK23', 0, 3)
    spreadsheet_redactor.prepare_setBorder_bot('A24:AK24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_bot('B1:AK1', 0, 2)
    spreadsheet_redactor.prepare_setBorder_bot('B2:AK2', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('A1:A24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('D1:D24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('G1:G24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('J1:J24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('M1:M24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('P1:P24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('S1:S24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('V1:V24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('Y1:Y24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('AB1:AB24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('AE1:AE24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('AH1:AH24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('AK1:AK24', 0, 3)
    spreadsheet_redactor.prepare_setBorder_right('B2:B24', 0)
    spreadsheet_redactor.prepare_setBorder_right('C2:C24', 0)
    spreadsheet_redactor.prepare_setBorder_right('E2:E24', 0)
    spreadsheet_redactor.prepare_setBorder_right('F2:F24', 0)
    spreadsheet_redactor.prepare_setBorder_right('H2:H24', 0)
    spreadsheet_redactor.prepare_setBorder_right('I2:I24', 0)
    spreadsheet_redactor.prepare_setBorder_right('K2:K24', 0)
    spreadsheet_redactor.prepare_setBorder_right('L2:L24', 0)
    spreadsheet_redactor.prepare_setBorder_right('N2:N24', 0)
    spreadsheet_redactor.prepare_setBorder_right('O2:O24', 0)
    spreadsheet_redactor.prepare_setBorder_right('Q2:Q24', 0)
    spreadsheet_redactor.prepare_setBorder_right('R2:R24', 0)
    spreadsheet_redactor.prepare_setBorder_right('T2:T24', 0)
    spreadsheet_redactor.prepare_setBorder_right('U2:U24', 0)
    spreadsheet_redactor.prepare_setBorder_right('W2:X24', 0)
    spreadsheet_redactor.prepare_setBorder_right('Z2:Z24', 0)
    spreadsheet_redactor.prepare_setBorder_right('AA2:AA24', 0)
    spreadsheet_redactor.prepare_setBorder_right('AC2:AC24', 0)
    spreadsheet_redactor.prepare_setBorder_right('AD2:AD24', 0)
    spreadsheet_redactor.prepare_setBorder_right('AF2:AG24', 0)
    spreadsheet_redactor.prepare_setBorder_right('AI2:AI24', 0)
    spreadsheet_redactor.prepare_setBorder_right('AJ2:AJ24', 0)
    spreadsheet_redactor.runPrepared()
    columns_list = ('B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L')
    for sheet in range(1, 13):
        if sheet in (1, 3, 5, 7, 8, 10, 12,):
            spreadsheet_redactor.prepare_setValues('A3:A33', [
                ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18',
                 '19', '20', '21', '22', '23', '24', '25', '26',
                 '27', '28', '29', '30', '31', ]], sheet, 'COLUMNS')
            spreadsheet_redactor.prepare_setValues('D34:D34', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setValues('G34:G34', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setValues('H34:H34', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setBorder_left('B1:B33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('A32:A33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('A33:L33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('D34:D35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('G34:G35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('H34:H35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('D2:D35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('G2:G35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('H2:H35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('J2:J33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('L2:L33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('B2:B33', sheet)
            spreadsheet_redactor.prepare_setBorder_right('C2:C33', sheet)
            spreadsheet_redactor.prepare_setBorder_right('E2:E33', sheet)
            spreadsheet_redactor.prepare_setBorder_right('F2:F33', sheet)
            spreadsheet_redactor.prepare_setBorder_right('I2:I33', sheet)
            spreadsheet_redactor.prepare_setBorder_right('K2:K33', sheet)
            spreadsheet_redactor.prepare_setBorder_right('C34:C35', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('F34:F35', sheet, 3)
            for index, col in enumerate(columns_list):
                if col in ('D', 'G', 'H',):
                    spreadsheet_redactor.prepare_setValues(f'{col}35:{col}35', [[f'=SUM({col}3:{col}33)']], sheet)
                if col in ('D', 'G',):
                    for row in range(3, 34):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}',
                                                               [[
                                                                   f'=SUM({columns_list[index - 2]}{row}:{columns_list[index - 1]}{row})']],
                                                               sheet)
                if col in ('H'):
                    for row in range(3, 34):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=D{row}-G{row}']],
                                                               sheet)
                if col in ('I'):
                    for row in range(3, 34):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=B{row}-E{row}']],
                                                               sheet)
                if col in ('J'):
                    for row in range(3, 34):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=C{row}-F{row}']],
                                                               sheet)

        elif sheet in (4, 6, 9, 11,):
            spreadsheet_redactor.prepare_setValues('A3:A32', [
                ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18',
                 '19', '20', '21', '22', '23', '24', '25', '26',
                 '27', '28', '29', '30', ]], sheet, 'COLUMNS')
            spreadsheet_redactor.prepare_setValues('D33:D33', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setValues('G33:G33', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setValues('H33:H33', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setBorder_left('B1:B32', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('A31:A32', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('A32:L32', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('D33:D34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('G33:G34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('H33:H34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('D2:D34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('G2:G34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('H2:H34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('J2:J32', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('L2:L32', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('B2:B32', sheet)
            spreadsheet_redactor.prepare_setBorder_right('C2:C32', sheet)
            spreadsheet_redactor.prepare_setBorder_right('E2:E32', sheet)
            spreadsheet_redactor.prepare_setBorder_right('F2:F32', sheet)
            spreadsheet_redactor.prepare_setBorder_right('I2:I32', sheet)
            spreadsheet_redactor.prepare_setBorder_right('K2:K32', sheet)
            spreadsheet_redactor.prepare_setBorder_right('C33:C34', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('F33:F34', sheet, 3)
            for index, col in enumerate(columns_list):
                if col in ('D', 'G', 'H',):
                    spreadsheet_redactor.prepare_setValues(f'{col}34:{col}34', [[f'=SUM({col}3:{col}32)']], sheet)
                if col in ('D', 'G',):
                    for row in range(3, 33):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}',
                                                               [[
                                                                   f'=SUM({columns_list[index - 2]}{row}:{columns_list[index - 1]}{row})']],
                                                               sheet)
                if col in ('H'):
                    for row in range(3, 33):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=D{row}-G{row}']],
                                                               sheet)
                if col in ('I'):
                    for row in range(3, 33):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=B{row}-E{row}']],
                                                               sheet)
                if col in ('J'):
                    for row in range(3, 33):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=C{row}-F{row}']],
                                                               sheet)

        elif sheet == 2:
            spreadsheet_redactor.prepare_setValues('A3:A31', [
                ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18',
                 '19', '20', '21', '22', '23', '24', '25', '26',
                 '27', '28', '29', ]], sheet, 'COLUMNS')
            spreadsheet_redactor.prepare_setValues('D32:D32', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setValues('G32:G32', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setValues('H32:H32', [['За месяц', ]], sheet)
            spreadsheet_redactor.prepare_setBorder_left('B1:B31', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('A30:A31', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('A31:L31', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('D32:D33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('G32:G33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_bot('H32:H33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('D2:D33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('G2:G33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('H2:H33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('J2:J31', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('L2:L31', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('B2:B31', sheet)
            spreadsheet_redactor.prepare_setBorder_right('C2:C31', sheet)
            spreadsheet_redactor.prepare_setBorder_right('E2:E31', sheet)
            spreadsheet_redactor.prepare_setBorder_right('F2:F31', sheet)
            spreadsheet_redactor.prepare_setBorder_right('I2:I31', sheet)
            spreadsheet_redactor.prepare_setBorder_right('K2:K31', sheet)
            spreadsheet_redactor.prepare_setBorder_right('C32:C33', sheet, 3)
            spreadsheet_redactor.prepare_setBorder_right('F32:F33', sheet, 3)
            for index, col in enumerate(columns_list):
                if col in ('D', 'G', 'H',):
                    spreadsheet_redactor.prepare_setValues(f'{col}33:{col}33', [[f'=SUM({col}3:{col}31)']], sheet)
                if col in ('D', 'G',):
                    for row in range(3, 32):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}',
                                                               [[
                                                                   f'=SUM({columns_list[index - 2]}{row}:{columns_list[index - 1]}{row})']],
                                                               sheet)
                if col in ('H'):
                    for row in range(3, 32):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=D{row}-G{row}']],
                                                               sheet)
                if col in ('I'):
                    for row in range(3, 32):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=B{row}-E{row}']],
                                                               sheet)
                if col in ('J'):
                    for row in range(3, 32):
                        spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}', [[f'=C{row}-F{row}']],
                                                               sheet)

        spreadsheet_redactor.prepare_setValues('B1:C1', [['Получено', ]], sheet)
        spreadsheet_redactor.prepare_setValues('E1:F1', [['Потрачено', ]], sheet)
        spreadsheet_redactor.prepare_setValues('H1:H1', [['Итого', ]], sheet)
        spreadsheet_redactor.prepare_setValues('I1:I1', [['На счетах', ]], sheet)
        spreadsheet_redactor.prepare_setValues('K1:K1', [['Комментарий', ]], sheet)
        spreadsheet_redactor.prepare_setValues('B2:G2', [[user_1, user_2, 'Сумма', user_1, user_2, 'Сумма', ]], sheet)
        spreadsheet_redactor.prepare_setValues('I2:J2', [[user_1, user_2]], sheet)
        spreadsheet_redactor.prepare_setValues('K2:L2', [[user_1, user_2]], sheet)
        spreadsheet_redactor.prepare_setCellsFormat('A1:A33',
                                                    {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE',
                                                     'textFormat': {'bold': True}}, sheet)
        spreadsheet_redactor.prepare_setCellsFormat('B1:L2',
                                                    {'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE',
                                                     'textFormat': {'bold': True}}, sheet)
        spreadsheet_redactor.prepare_setValues('A1:A1', [['Дата']], sheet)
        spreadsheet_redactor.prepare_setBorder_bot('A1:A2', sheet, 3)
        spreadsheet_redactor.prepare_setBorder_bot('H1:H2', sheet, 3)
        spreadsheet_redactor.prepare_setBorder_bot('B1:G1', sheet, 2)
        spreadsheet_redactor.prepare_setBorder_bot('I1:L1', sheet, 2)
        spreadsheet_redactor.prepare_setBorder_bot('B2:L2', sheet, 3)
        spreadsheet_redactor.runPrepared()
        columns_list = (
            'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V',
            'W',
            'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK',)
        for index, col in enumerate(columns_list):
            spreadsheet_redactor.prepare_setValues(f'{col}24:{col}24', [[f'=SUM({col}3:{col}23)']], 0)
            if col in ('D', 'G', 'J', 'M', 'P', 'S', 'V', 'Y', 'AB', 'AE', 'AH', 'AK',):
                for row in range(3, 24):
                    spreadsheet_redactor.prepare_setValues(f'{col}{row}:{col}{row}',
                                                           [[
                                                               f'=SUM({columns_list[index - 2]}{row}:{columns_list[index - 1]}{row})']],
                                                           0)
        spreadsheet_redactor.runPrepared()


def define_category(data):
    profit_categories = (
        'зп', 'зарплата', 'кэшбек', 'кешбек', 'возврат', 'пенсия', 'стипендия', 'подарили', 'проценты', 'пособие',
        'выигрыш', 'приз', 'от родителей', 'от детей', 'аванс', 'возврат налогов', 'отпускные',
        'командировочные',)
    loss_categories = {'Продукты': ('продукты', 'окей',),
                       'Красота': ('стрижка', 'косметика', 'подружка', 'бритье',),
                       'Медицина': ('стоматолог', 'анализы', 'справка',),
                       'Аптека': ('аптека', 'лекарства',),
                       'Одежда': ('одежда',),
                       'Развлечения': ('развлечения',),
                       'Спорт': ('спорт', 'сноуборд', 'лыжи', 'бег'),
                       'Связь': (
                           'связь', 'интернет', 'мтс', 'домру', 'ростелеком', 'мегафон', 'билайн', 'теледва', 'йота',
                           'интерзет', 'interzet',),
                       'Транспорт': (
                           'метро', 'такси', 'трамвай', 'автобус', 'тролейбус', 'самолет', 'бензин', 'транспорт',
                           'каршеринг'),
                       'Услуги банка': ('услуги банка', 'смс оповещения',),
                       'Домашние животные': ('кошка', 'собака', 'домашние животные',),
                       'ЖКХ': ('жкз', 'квартплата',),
                       'Сервисы(музыка/сериалы/игры)': (
                           'айтюнс', 'яндекс музыка', 'гугл музыка', 'itunes', 'музыка', 'подписка', 'нетфликс',
                           'netflix', 'hbo',
                           'amediateka', 'okko', 'ivi', 'steam', 'egs', 'игра', 'origin', 'плойка', 'play station',),
                       'Кафе и рестораны': (
                           'кафе', 'ресторан', 'бар', 'макдак', 'кфс', 'kfc', 'мак', 'кофе', 'пиво', 'кабак',),
                       'Переводы': ('перевод',),
                       'Ипотека/Аренда': ('ипотека', 'аренда',),
                       'Техника': ('техника',),
                       'Ремонт': ('ремонт',),
                       'Для дома': ('для дома', 'бытовая химия',),
                       'Подарки': ('подарок', 'подарки'),
                       'Разное': ('разное',
                                  '',), }  # пустая строка '' здесь должна быть. если запись вносится в таблицу просто суммой, без категории, она пишется сюда
    current_category = ''
    category_row = None
    if data.lower() in profit_categories:  # неправильно
        profit = True
    else:
        for category, _ in loss_categories.items():
            if data.lower() in _:
                profit = False
                current_category = category
        for index, category in enumerate(loss_categories):
            if current_category == category:
                category_row = index + 3
    return profit, current_category, category_row


if __name__ == '__main__':
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    SERVICE_ACCOUNT_FILE = 'Kekik-cdb630d21d71.json'
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service_for_drive = build('drive', 'v3', credentials=credentials)
    service_for_sheets = build('sheets', 'v4', credentials=credentials)
    # user_2 = gmail пользователь дает боту свой
    # user_1 = gmail
    # user_2_name = парсится чат с ботом и достается имя
    # user_1_name =
    spreadsheet = create_spreadsheet()  # Создаем таблицу
    permissions_for_owner(user_2)  # Передаем права на владение
    # permissions_for_second_user(user_2) # Даем права на редактирования для второго пользователя
    spreadsheet_id = get_spreadsheet_id()  # Получаем ID
    spreadsheet_redactor = Spreadsheet()  # Создаем экземпляр класса для редактирования таблицы
    prepare_sheets(spreadsheet_redactor)  # Длеаем слияние ячеек для заполнения в соответствии с шаблоном
    create_template(spreadsheet_redactor)  # Создаем шаблон, который будем наполнять своими данными
    # spreadsheet_redactor.update_cell_data(data='услуги банка 1000', date=1576000479.749238, user_name=user_1_name)
