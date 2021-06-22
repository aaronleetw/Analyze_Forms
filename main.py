from Google import Create_Service
import csv
from datetime import datetime
import os
from time import sleep
from tqdm import tqdm

CLIENT_SECRET_FILE = 'client_secret.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


service = Create_Service(
    CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

SPREADSHEET_IDS_LIST = [
    '1aD3sMh1Y4DMWccvAKFNpMhS2rjMYRfR7rcBAYff7x_Y',
    '1UYPNVTcFap_QwEgV45HN0X8mn0QKVqkxI85tnXh_vkc'
]
CLASS = ["ping", "ho"]


for cla in range(2):
    SPREADSHEET_IDS = []
    SCORES = []
    AUTHOR_SCORES = []
    QUESTION_ANSWER = []
    i = 1
    while (i <= 33):
        SCORES.append([i, 0, 0])
        AUTHOR_SCORES.append([i, 0, 0])
        QUESTION_ANSWER.append([i, 0])
        i += 1
    del i

    OUTPUT_DIR = os.path.join(os.getcwd(), 'output_' + CLASS[cla])
    LEGAL_TARGET = os.path.join(OUTPUT_DIR, 'legal.csv')
    ILLEGAL_TARGET = os.path.join(OUTPUT_DIR, 'illegal.csv')
    ANALYZED_SCORES_TARGET = os.path.join(
        OUTPUT_DIR, 'analyzed_scores.csv')
    ANALYZED_AUTHOR_TARGET = os.path.join(
        OUTPUT_DIR, 'analyzed_author.csv')
    QA_TARGET = os.path.join(OUTPUT_DIR, 'qa.csv')

    open(LEGAL_TARGET, 'w').write("Timestamp,Author,Number,Name,Score\n")
    open(ILLEGAL_TARGET, 'w').write("Timestamp,Author,Number,Name,Score\n")
    open(ANALYZED_SCORES_TARGET, 'w').write(
        "Number,Score(/50),ScoredCount,Author1,Score1,Author2,Score2,Author3,Score3,Author4,Score4,Author5,Score5\n")
    open(ANALYZED_AUTHOR_TARGET, 'w').write(
        "Author,AvgScore(/100),ScoredCount,FilledStud1,Score1,FilledStud2,Score2,FilledStud3,Score3,FilledStud4,Score4,FilledStud5,Score5\n")
    open(QA_TARGET, 'w').write(
        "Author,Code,Q1,Q2,Q3,Q4,Q5,Q6,Q7,Q8,Q9,Q10,A1,A2,A3,A4,A5,A6,A7,A8,A9,A10\n"
    )

    print("GETTING SPREADSHEET_IDS (" + CLASS[cla] + ")")
    response = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_IDS_LIST[cla],
        majorDimension='ROWS',
        range='B:C'
    ).execute()
    for r in response['values'][1:]:
        SPREADSHEET_IDS.append(r)
    del response
    print(".....DONE")
    print("START PROCESSING " + str(len(SPREADSHEET_IDS)) + " spreadsheets")

    for id in SPREADSHEET_IDS:
        response = ""
        afternoon = False
        try:
            response = service.spreadsheets().values().get(
                spreadsheetId=id[1],
                majorDimension='ROWS',
                range='A:N'
            ).execute()
        except:
            print("{0:2d}".format(int(id[0])), end=' ')
            print('raised an error(' + str(id[1]) + ')')
            continue
        if (QUESTION_ANSWER[int(id[0]) - 1][1] == 0):
            QUESTION_ANSWER[int(id[0]) - 1].extend(response['values'][0][4:])
            QUESTION_ANSWER[int(id[0]) - 1][1] += 1
        for r in response['values'][1:]:
            r.insert(1, id[0])
            r[2] = r[2].removesuffix(' / 100')
            r += [r.pop(2)]
            for i in range(10):
                r += [r.pop(4)]

            if (int(r[4]) == 100 and QUESTION_ANSWER[int(id[0])-1][1] == 1):
                QUESTION_ANSWER[int(id[0]) - 1].extend(r[5:])
                QUESTION_ANSWER[int(id[0]) - 1][1] += 1

            submitDateTime = datetime.now
            try:
                submitDateTime = datetime.strptime(r[0], '%m/%d/%Y %H:%M:%S')
            except:
                try:
                    submitDateTime = datetime.strptime(
                        r[0], '%Y-%m-%d %H:%M:%S')
                except:
                    try:
                        submitDateTime = datetime.strptime(
                            r[0], '%Y/%m/%d 下午 %H:%M:%S')
                        afternoon = True
                    except:
                        try:
                            submitDateTime = datetime.strptime(
                                r[0], '%Y/%m/%d 上午 %H:%M:%S')
                        except:
                            try:
                                submitDateTime = datetime.strptime(
                                    r[0], '%Y-%m-%d 下午%H:%M:%S')
                                afternoon = True
                            except:
                                try:
                                    submitDateTime = datetime.strptime(
                                        r[0], '%Y-%m-%d 上午%H:%M:%S')
                                except:
                                    print("{0:2d}".format(int(id[0])), end=' ')
                                    print(
                                        'raised a TIME error(' + str(id[1]) + ')')
            if (afternoon == True):
                submitDateTime.replace(hour=submitDateTime.hour + 12)
            r[0] = submitDateTime.strftime("%m/%d/%Y %H:%M:%S")
            if not (((submitDateTime.day == 22 and submitDateTime.hour >= 12) or
                    submitDateTime.day == 23 or
                    (submitDateTime.day == 24 and submitDateTime.hour < 12)) and
                    submitDateTime.year == 2021 and submitDateTime.month == 6):
                with open(ILLEGAL_TARGET, 'a', encoding='UTF8', newline='') as f:
                    csv.writer(f).writerow(r[:5])
                continue
            with open(LEGAL_TARGET, 'a', encoding='UTF8', newline='') as f:
                csv.writer(f).writerow(r[:5])
            SCORES[int(r[2]) - 1][2] += 1
            SCORES[int(r[2]) - 1][1] += int(int(r[4]) * 0.1)
            SCORES[int(r[2]) - 1].append(id[0])
            SCORES[int(r[2]) - 1].append(int(r[4]))
            AUTHOR_SCORES[int(id[0]) - 1][2] += 1
            AUTHOR_SCORES[int(id[0]) - 1][1] += int(r[4])
            AUTHOR_SCORES[int(id[0]) - 1].append(r[2])
            AUTHOR_SCORES[int(id[0]) - 1].append(int(r[4]))
        del response
    for s in SCORES:
        with open(ANALYZED_SCORES_TARGET, 'a', encoding='UTF8', newline='') as f:
            csv.writer(f).writerow(s)
    for s in AUTHOR_SCORES:
        if (s[2] != 0):
            s[1] /= s[2]
            s[1] = int(s[1])
        with open(ANALYZED_AUTHOR_TARGET, 'a', encoding='UTF8', newline='') as f:
            csv.writer(f).writerow(s)
    for qa in QUESTION_ANSWER:
        with open(QA_TARGET, 'a', encoding='UTF8', newline='') as f:
            csv.writer(f).writerow(qa)
    print(".....DONE\n")
    if (cla != 1):
        for xu in tqdm(range(20)):
            sleep(3)
