from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly','https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID = None
SAMPLE_RANGE_NAME = "'Formresponses1'"

def obtaincreds():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def data():
    creds = obtaincreds()

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    return(values)

def checksold(number):
    RANGE = SAMPLE_RANGE_NAME + '!A' + str(number) + ':A' + str(number)
    creds = obtaincreds()

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    result = sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID, ranges=RANGE, includeGridData=True).execute()
    resultcolor = result['sheets'][0]['data'][0]['rowData'][0]['values'][0]['effectiveFormat']['backgroundColor']
    if resultcolor == {'red': 1, 'green': 1, 'blue': 1}:
        return False
    else:
        return True

def sell(index):
    RANGE = SAMPLE_RANGE_NAME + '!' + str(index) + ':' + str(index)
    creds = obtaincreds()
    service = build('sheets', 'v4', credentials=creds)
    requests = [
    {'updateCells': {
        'rows':[{'values': {'userEnteredFormat': {'backgroundColor': {'red': 0, 'blue': 0, 'green': 1}}}}],
        'range': {
            'sheetId': 67405509,
            'startRowIndex': index-1,
            'endRowIndex': index,
        },
        'fields': 'userEnteredFormat.backgroundColor'
    }}
    ]
    body = {
        'requests': requests
    }
    sheet = service.spreadsheets()
    result = sheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()

def updateinfo(vo, info):
    index = vo.index
    RANGE = SAMPLE_RANGE_NAME + '!Z' + str(index) + ':Z' + str(index)
    body = {'range': RANGE, 'values': [[info]]}
    creds = obtaincreds()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=RANGE, valueInputOption='RAW', body=body).execute()


class TempAccessObject:
    def __init__(self, processobject):
        _tempobject = processobject[0]
        self.sold = processobject[2]
        self.index = processobject[1]
        self.name = _tempobject[2]
        self.clas = _tempobject[3]
        self.email = _tempobject[1]
        self.phoneno = _tempobject[24]
        self._raworderdata = _tempobject[5:22]
        self.sizes = ['XXS', 'XS', 'S', 'M', 'L', 'XL']
        self.designs = ['Green', 'Black', 'Singlet']
        self.order = {}
        iterator = 0
        for p in self.designs:
            self.order[p] = {}
            for r in self.sizes:
                try:
                    self.order[p][r] = int(self._raworderdata[iterator])
                except:
                    self.order[p][r] = 0
                iterator += 1

def updateorder(vo):
    updatedlist = list()
    for a in vo.order.keys():
        for b in vo.order[a].keys():
            if vo.order[a][b] != 0:
                updatedlist.append(str(vo.order[a][b]))
            else:
                updatedlist.append('')
    index = vo.index
    RANGE = SAMPLE_RANGE_NAME + '!F' + str(index) + ':W' + str(index)
    body = {'range': RANGE, 'values': [updatedlist]}
    creds = obtaincreds()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=RANGE, valueInputOption='RAW', body=body).execute()

def editorder(vo):
    editloop = True
    while editloop:
        print('_'*10)
        print('Welcome to the edit screen!')
        print('Current Orders:')
        for j in vo.order.keys():
            for f in vo.order[j].keys():
                if vo.order[j][f] != 0:
                    print(str(vo.order[j][f]) + ' ' + f + ' ' + j)
        editcommand = input('Enter edit command:')
        if editcommand == '':
            print('Invalid Command!')
        elif editcommand[0] == '+':
            try:
                design = int(editcommand[-1]) - 1
                sizes = int(editcommand[-2]) - 1
                editcommand = editcommand[:-2]
                editcommand = editcommand[1:]
                numberofshirts = int(editcommand)
                temporder = vo.order
                tempsizeorder = temporder[vo.designs[design]]
                tempsizeorder[vo.sizes[sizes]] = numberofshirts
                temporder[vo.designs[design]] = tempsizeorder
                vo.order = temporder
            except:
                print('Invalid Command!')
        elif editcommand.lower() in ['f','finish']:
            updateorder(vo)
            editloop = False

if __name__ == '__main__':
    obtaincreds()
    while True:
        print('_'*10)
        print('Welcome to the application!')
        print('_'*10)
        c = True
        phoneselect = False
        while c:
            namesearch = input('Search Name:')
            if namesearch == '':
                print('Search cannot be blank!')
            elif namesearch.lower() in ['phonesearch', 'phone search']:
                g = True
                while g:
                    phonesearch = input('Search Phone:')
                    if phonesearch == '':
                        print('Search cannot be blank!')
                    else:
                        g = False
                        phoneselect = True
                        c = False
            else:
                c = False
        spreadsht = data()
        title = spreadsht.pop(0)
        selectlst = list()
        count = 2
        if not phoneselect:
            for x in spreadsht:
                if namesearch.lower() in x[2].lower():
                    while len(x) < 26:
                        x.append('')
                    infolist = [x, count]
                    selectlst.append(infolist)
                count += 1
        else:
            for x in spreadsht:
                if len(x) > 24:
                    if phonesearch in x[24]:
                        while len(x) < 26:
                            x.append('')
                        infolist = [x, count]
                        selectlst.append(infolist)
                count +=1
        templist = list()
        if len(selectlst) < 6:
            sold_checked = True
            for x in selectlst:
                t = x
                t.append(checksold(x[1]))
                templist.append(t)
        else:
            sold_checked = False
            print('List too long, operation too inefficient, unable to check for sold status.')
            for x in selectlst:
                t = x
                t.append(False)
                templist.append(t)
        selectlst = templist
        if len(selectlst) == 0:
            print('No results found!')
        else:
            counter = 1
            for i in selectlst:
                if i[2]:
                    soldstatus = '(SOLD)'
                else:
                    soldstatus = ''
                print(str(counter)+'. '+i[0][2] + ', ' + i[0][3] + soldstatus)
                counter +=1
            r = True
            while r:
                noselect = input('Select choice:')
                try:
                    info = selectlst[int(noselect)-1]
                    r = False
                except:
                    print('Invalid Entry!')
            vo = TempAccessObject(info)
            if not sold_checked:
                vo.sold = checksold(vo.index)
            UIloop = True
            while UIloop:
                print('_'*10)
                print('Name:' + vo.name)
                print('Email:' + vo.email)
                print('Class:' + vo.clas)
                print('Phone:' + vo.phoneno)
                print('Here are the orders:')
                for j in vo.order.keys():
                    for f in vo.order[j].keys():
                        if vo.order[j][f] != 0:
                            print(str(vo.order[j][f]) + ' ' + f + ' ' + j)
                if vo.sold:
                    print('THIS ITEM IS SOLD! DO NOT SELL!')
                cmd = input('Command:')
                print('_'*10)
                if cmd.lower() in ['sell', 's']:
                    if not vo.sold:
                        sell(vo.index)
                        print('~Item has sold!~')
                    else:
                        print("You can't sell the same order twice!")
                    UIloop = False
                elif cmd.lower() in ['sell*', 's*']:
                    if not vo.sold:
                        information = input('Enter Contact Info:')
                        sell(vo.index)
                        updateinfo(vo, information)
                        print('~Item has sold!~')
                        UIloop = False
                    else:
                        print("You can't sell the same order twice!")
                        UIloop = False
                elif cmd.lower() == '+':
                    editorder(vo)
                else:
                    UIloop = False
        input('Press enter to continue...')
