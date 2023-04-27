import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

main_dict = {}

CREDENTIALS_FILE = 'creds.json'
spreadsheet_id = '1Nz_VX7fYU2O_TuM2XtmIRHlvmJofUkD5BnR5tkyw5so'
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

participation_stage_dict = {
    'passed_registration': 'Прошёл регистрацию',
    'wrote_qualifying': 'Написал отборочный этап',
    'passed_final': 'Прошёл на заключительный этап',
    'took_final': 'Принял участие в финале(участник)',
    'final_prize_winner': 'Призёр финала(диплом 2 или 3 степени)',
    'winner_of_final': 'Победитель финала(диплом 1 степени',
}

mentors_dict = {

}

lessons_dict = {

}

olympiad_dict = {}  # Словарь, где ключ это индекс олимпиады, а значение это название олимпиады
olympiad_lst = []  # Список всех олимпиад


def table(row, range, regim, values=[]): # функция для чтения или записи в таблицу
    if regim == "read":
        values = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range="Лист1!" + range,
            majorDimension=row
        ).execute()
        if "values" in values:
            num = values["values"]
            return num
    else:
        results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": "Лист1!" + range,
                 "majorDimension": row,
                 "values": values},
            ]
        }).execute()


def olympSheet():
    olympiad_dict.clear()
    olympiad_lst.clear()
    olympiad = table("ROWS", "A1:B250", "read")
    for i in olympiad:
        if len(i[0]) == 1:
            i[0] = "0" + i[0]
        if i[0] not in olympiad_dict:
            olympiad_dict[i[0]] = [i[1]]
        else:
            olympiad_dict[i[0]].append([i[1]])
    olympiad = sorted(olympiad, key=lambda x: x[1])
    string = ''
    c = 0
    for i in olympiad:
        if "НТО:" not in i[1]:
            i[0] += ":"
        if c <= 15:
            string += " ".join(i)
            c += 1
        else:
            olympiad_lst.append(string)
            string, c = " ".join(i), 0
        string += "\n" + " " + "\n"
    olympiad_lst.append(string)


def main2():
    lessons_dict.clear()
    mentors_dict.clear()
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="Лист2!A1:A100",
        majorDimension="COLUMNS"
    ).execute()
    values = values["values"][0]
    for i in range(len(values)):
        mentors_dict['mentor' + str(i + 1)] = values[i]

    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="Лист3!A1:A100",
        majorDimension="COLUMNS"
    ).execute()
    values = values["values"][0]
    for i in range(len(values)):
        lessons_dict['lesson' + str(i + 1)] = values[i]
