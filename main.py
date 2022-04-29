from jira import JIRA
import iso8601
import time
import datetime
import openpyxl
import sys


def excel_time_convert(tz_datetime):
    d = iso8601.parse_date(tz_datetime)
    d = d.astimezone(None)
    return datetime.datetime.combine(d.date(), d.time(), None)
    # "{:0>2}.{:0>2}.{} {}".format(d.day, d.month, d.year, d.time())


def as_str(v, default='', attr0=None, level=0):
    if isinstance(v, list) and level == 0:
        return ', '.join([as_str(x, default, attr0, 1) for x in v])
    if v is None:
        return default
    if attr0 is None:
        return '{}'.format(v)
    return '{}'.format(v.__getattribute__(attr0))


if __name__ == '__main__':
    wb = openpyxl.workbook.Workbook()
    ws = wb.active
    
    try: 
        host = sys.argv[1]
        print(f'host: {host}')
        user_name = sys.argv[2]
        print(f'user_name: {user_name}')
        pwd = sys.argv[3]
        if len(pwd) > 0: 
            print('pwd: ********')
        file_path = sys.argv[4]
        print(f'file_path {file_path}')
        jql = sys.argv[5]
        print(f'jql: {jql}')
        jira_options = {'server': host}
        jira = JIRA(options=jira_options, basic_auth=(user_name, pwd))
        del pwd
    except:
        print('Incorrect parameters');
        print(f'Usage: {sys.argv[0]} <jira server URL> <user_name> <password> <excel file> <Jira Query>')
        raise


    def get_field(field_name):
        for f in jira.fields():
            if f["name"] == field_name:
                return f["id"]
        return None


    def get_value(issue, field):
        f = get_field(field)
        if f is None:
            return None
        return issue.fields.__getattribute__(f)

    print('getting the list of jira tasks, max result 300')
    
    issues_list = jira.search_issues(jql, maxResults=300)

    print('success')

    row0 = ['project', 'type', 'key', 'summary', 'priority', 'status', 'created', 'updated', 'reporter', 'assignee',
            'External Suppliers Involved', 'Sprint assigned', 'Wave Organizing indicator']
    ws.append(row0)

    print('collecting data')

    for issue in issues_list:
        ws.append(
            [
                as_str(issue.fields.project, attr0='name'),
                as_str(issue.fields.issuetype, attr0='name'),
                issue.key,
                issue.fields.summary,
                as_str(issue.fields.priority, attr0='name'),
                as_str(issue.fields.status, attr0='name'),
                excel_time_convert(issue.fields.created),
                excel_time_convert(issue.fields.updated),
                as_str(issue.fields.creator, default='Unknown', attr0='displayName'),
                as_str(issue.fields.assignee, default='Unassigned', attr0='displayName'),
                as_str(get_value(issue, 'External Suppliers Involved'), attr0='value'),
                as_str(get_value(issue, 'Sprint assigned'), attr0='value'),
                as_str(get_value(issue, 'Wave Organizing indicator'), attr0='value')
            ])

    print('saving...')

    try:
        wb.save(file_path)
        wb.close()
    except:
        print(f'Exception {sys.exc_info()[0]}')

    print('exiting in 3 seconds...')

    time.sleep(3)
