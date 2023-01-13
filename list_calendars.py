from cal_setup import get_calendar_service

def list_of_calendars():
    service = get_calendar_service()
    # Call the Calendar API
    print('Sto prendendo la lista dei calendari')
    calendars_result = service.calendarList().list().execute()

    calendars = calendars_result.get('items', [])

    if not calendars:
       print('Nessun calendario trovato')
    else:
        return calendars
    for calendar in calendars:
       summary = calendar['summary']
       id = calendar['id']
       primary = "Primary" if calendar.get('primary') else ""
       print("%s\t%s\t%s" % (summary, id, primary))