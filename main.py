import sys
from datetime import timedelta, datetime

from dateutil.relativedelta import relativedelta
from oauth2client import client
from googleapiclient import sample_tools
from decouple import config


def main(argv):
    # Authenticate and construct service.
    calendar, flags = sample_tools.init(
        argv, 'calendar', 'v3', __doc__, __file__,
        scope='https://www.googleapis.com/auth/calendar')

    sheets, flags = sample_tools.init(
        argv, 'sheets', 'v4', __doc__, __file__,
        scope='https://www.googleapis.com/auth/spreadsheets')

    try:
        today = datetime.today()
        day, month, year = today.day, today.month, today.year
        first_day_of_month = datetime(year, month, 1).isoformat() + 'Z'
        last_day_of_month = ((datetime(year, month, 1) + relativedelta(months=1)) - timedelta(days=1)).isoformat() + 'Z'
        today = datetime.today().isoformat() + 'Z'
        events_result = calendar.events().list(calendarId='primary', timeMin=today, timeMax=last_day_of_month,
                                               singleEvents=True, orderBy='startTime').execute()
        events = events_result.get('items', [])

        spreadsheet_id = config('SPREADSHEET')
        range_name = config('RANGE')

        values = []
        for event in events:
            nome_evento = event['summary']
            if 'MR' in nome_evento:
                # Manutenção Russo
                values.append([event['summary'], 90])
            if 'MB' in nome_evento:
                # Manutenção Brasileiro
                values.append([event['summary'], 70])
            if 'AR' in nome_evento:
                # Aplicação Russo
                values.append([event['summary'], 150])
            if 'AB' in nome_evento:
                # Aplicação Brasileiro
                values.append([event['summary'], 120])
            if 'SO' in nome_evento:
                # Sobrancelha
                values.append([event['summary'], 20])
            if 'PT' in nome_evento:
                # Penteado na Taina
                values.append([event['summary'], 84])

        body = {
            'values': values
        }

        result = sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption="USER_ENTERED", body=body).execute()
        print('{0} cells updated.'.format(result.get('updatedCells')))


    except client.AccessTokenRefreshError:
        print('The credentials have been revoked or expired, please re-run'
              'the application to re-authorize.')


if __name__ == '__main__':
    main(sys.argv)
