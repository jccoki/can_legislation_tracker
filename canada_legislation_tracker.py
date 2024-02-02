#!/usr/bin/env python
import win32com.client as win32 #external
import dateparser #external
import yaml #external
import openpyxl #external
from selenium import webdriver #external
from selenium.webdriver.common.keys import Keys #external
from selenium.webdriver.common.by import By #external
from selenium.webdriver.support.wait import WebDriverWait #external
from selenium.webdriver.support import expected_conditions as EC #external
import calendar
import pytz
from webdriver_manager.chrome import ChromeDriverManager
import datetime

def generate_date_range(date1, date2):
    for d in range(int ((date2 - date1).days) + 1):
        yield date1 + datetime.timedelta(d)

# excludes weekends on business days counting
def add_business_days(start_date, add_days):
    current_date = start_date
    while add_days > 0:
        current_date = current_date + datetime.timedelta(days=1)
        weekday = current_date.isoweekday()
        if weekday >= 6:
            continue
        add_days = add_days - 1
    return current_date

# see chrome webdriver options at https://peter.sh/experiments/chromium-command-line-switches/ and
# https://www.selenium.dev/selenium/docs/api/py/webdriver_chrome/selenium.webdriver.chrome.webdriver.html
#chrome_options = webdriver.ChromeOptions()
#chrome_options.add_argument('--start-maximized')
#current_dir = Path.cwd()
#log_file = Path(current_dir, Path(__file__).stem + '.log')

# use a chrome webdriver manager as a default, then fallback on an local chrome webdriver
try:
    chrome_service = webdriver.ChromeService(ChromeDriverManager().install())
except:
    chrome_service = webdriver.ChromeService(executable_path = "resources/chromedriver.exe")
else:
    chrome_webdriver = webdriver.Chrome(service=chrome_service)
    chrome_webdriver.maximize_window()

# build date received schedule calenday dates
# reference material is the Link to Lexis.xlsx
date_received_schedule = {}

# can be part of config file
excel_date_format = "%m/%d/%Y"

config_file = open("config.yaml", 'r')
config_data = yaml.load(config_file, Loader=yaml.FullLoader)

# generate the row positions for ms excel values
excel_row_matrix = {}
jurisdiction_list = [
    'alberta', 'british columbia', 'federal english', 'federal french',
    'manitoba', 'new brunswick english', 'new brunswick french',
    'newfoundland', 'northwest territories', 'nova scotia', 'nunavut',
    'ontario', 'prince edward island', 'québec english', 'québec french',
    'saskatchewan', 'yukon']

# common legislation types for all jurisdiction
legislation_types = ['STAT Amendments', 'STAT New Doc', 'REG Amendments', 'REG New Doc']

# cell start row is part of a merged cell
start_row = 11
start_row_month = {}
for counter in range(1, 13):
    # generate month names
    month_name = calendar.month_name[counter].lower()

    # month are positioned on 80 row increments
    excel_row_matrix[month_name] = {}
    start_row_month[month_name] = start_row
    start_row = start_row + 80

    excel_dashboard_start_row = start_row_month[month_name] + 1
    for jurisdiction in jurisdiction_list:
        date_received_schedule[jurisdiction] = []

        excel_row_matrix[month_name][jurisdiction] = {}
        for legislation_type in legislation_types:
            excel_row_matrix[month_name][jurisdiction][legislation_type] = excel_dashboard_start_row
            excel_dashboard_start_row = excel_dashboard_start_row + 1

        # add annuals and sec legislation types for specific jurisdiction
        if (jurisdiction == 'alberta') or (jurisdiction == 'british columbia') or \
            (jurisdiction == 'federal english') or (jurisdiction == 'federal french') or \
            (jurisdiction == 'ontario'):
            excel_row_matrix[month_name][jurisdiction]['Annuals'] = excel_dashboard_start_row
            excel_dashboard_start_row = excel_dashboard_start_row + 1

        if jurisdiction == 'ontario':
            excel_row_matrix[month_name][jurisdiction]['SEC'] = excel_dashboard_start_row
            excel_dashboard_start_row = excel_dashboard_start_row + 1

# build date received schedule calendar dates
# reference material is the Link to Lexis.xlsx
year = config_data['Calendar Year']
#datetime.datetime.now().year
for month in range(1, 13):
    month_end = 30
    last_day = calendar.monthrange(year, month)[1]

    if month == 2:
        month_end = last_day

    date_received_schedule['alberta'].append(datetime.date(year, month, 15))
    date_received_schedule['alberta'].append(datetime.date(year, month, month_end))

    date_received_schedule['nova scotia'].append(datetime.date(year, month, 15))
    date_received_schedule['nova scotia'].append(datetime.date(year, month, month_end))

    for iterator in calendar.TextCalendar(calendar.FRIDAY).itermonthdays(year, month):
        if iterator != 0:
            day = datetime.date(year, month, iterator)
            # filter Fridays
            if day.isoweekday() == 5:
                date_received_schedule['british columbia'].append(day)
                date_received_schedule['manitoba'].append(day)

    for iterator in calendar.TextCalendar(calendar.WEDNESDAY).itermonthdays(year, month):
        if iterator != 0:
            day = datetime.date(year, month, iterator)
            # filter Wednesdays
            if day.isoweekday() == 3:
                date_received_schedule['federal english'].append(day)
                date_received_schedule['federal french'].append(day)

                date_received_schedule['new brunswick english'].append(day)
                date_received_schedule['new brunswick french'].append(day)

                date_received_schedule['québec english'].append(day)
                date_received_schedule['québec french'].append(day)

    for iterator in calendar.TextCalendar(calendar.SATURDAY).itermonthdays(year, month):
        if iterator != 0:
            day = datetime.date(year, month, iterator)
            # filter Saturdays
            if day.isoweekday() == 6:
                date_received_schedule['newfoundland'].append(day)
                date_received_schedule['ontario'].append(day)
                date_received_schedule['prince edward island'].append(day)
                date_received_schedule['saskatchewan'].append(day)

    date_received_schedule['northwest territories'].append(datetime.date(year, month, last_day))
    date_received_schedule['nunavut'].append(datetime.date(year, month, last_day))

    date_received_schedule['yukon'].append(datetime.date(year, month, 15))

lexisadvance_username = config_data['Quicklaw']['username']
lexisadvance_password = config_data['Quicklaw']['password']
lexisadvance_login_page = config_data['Quicklaw']['login page']
if (lexisadvance_username is None) or (lexisadvance_password is None) or \
    (lexisadvance_login_page is None):
    raise LookupError('Missing Lexis Advance Quicklaw login data')

currency_tracker_path = config_data['Currency Tracker']
if (currency_tracker_path is None):
    raise LookupError('Missing target Currency Tracker MS Excel file')

# convert into timezone aware date
utc = pytz.UTC
outlook_start_date = utc.localize(dateparser.parse(config_data['Outlook']['start date']))
outlook_end_date = utc.localize(dateparser.parse(config_data['Outlook']['end date']))
if (outlook_start_date is None) or (outlook_end_date is None):
    raise LookupError('Missing email date range in config file')

outlook_target_month = config_data['Outlook']['target month']
if outlook_target_month is None:
    raise LookupError('Missing target month designation in config file')

lexisadvance_links = config_data['Links']

# see https://github.com/hornlaszlomark/python_outlook
outlook = win32.Dispatch("Outlook.Application")
outlook_namespace = outlook.GetNamespace("MAPI")
outlook_account = config_data['Outlook']['account']
if (outlook_account is None):
    raise LookupError('Missing MS Outlook account data')

outlook_root_folder = outlook_namespace.Folders[outlook_account]
if outlook_root_folder is None:
    raise RuntimeError('Unable to access root folder under ' + outlook_account + ' account')
else:
    config_folder = config_data['Outlook']['folder']
    config_folder_parts = config_folder.split('/')
    outlook_target_folder = outlook_root_folder
    try:
        if len(config_folder_parts) == 1:
            outlook_target_folder = outlook_target_folder.Folders(config_folder_parts[0])
        else:
            for config_folder_part in config_folder_parts:
                outlook_target_folder = outlook_target_folder.Folders(config_folder_part)        
    except:
        print('Unable to locate the specified MS Outlook folder')
    else:
        outlook_messages = outlook_target_folder.Items
        # mail items are arranged from newest to oldest so we change the sort order using received time
        outlook_messages.Sort("[ReceivedTime]")

        # try to login first to ensure successive browser call        
        chrome_webdriver.get(lexisadvance_login_page)
        # we check the sign in button if we are facing a login face
        if chrome_webdriver.find_element(By.ID, "signInSbmtBtn") is not None:
            html_elem = chrome_webdriver.find_element(By.ID, "userid")
            html_elem.clear()
            html_elem.send_keys(lexisadvance_username)

            html_elem = chrome_webdriver.find_element(By.ID, "password")
            html_elem.clear()
            html_elem.send_keys(lexisadvance_password)
            
            chrome_webdriver.find_element(By.ID, "rememberMe").click()
            chrome_webdriver.find_element(By.ID, "signInSbmtBtn").click()
        
        try:
            excel_workbook = openpyxl.load_workbook(currency_tracker_path)
        except:
            print('Unable to open target currency tracker file')
        else:
            print('Getting list of Canada holidays: ' + str(currency_tracker_path))
            excel_worksheet_canada_holidays = excel_workbook['CAN Holidays']
            
            canada_holidays = []
            # holiday list starts on cell B5
            counter = 5
            canada_holiday = excel_worksheet_canada_holidays['B'+ str(counter)].value
            while canada_holiday:
                # value fetched from xlsx is cast to datetime object
                canada_holidays.append(canada_holiday.date())
                counter = counter + 1
                canada_holiday = excel_worksheet_canada_holidays['B'+ str(counter)].value

        for outlook_message in outlook_messages:
            excel_jurisdiction = ''
            excel_type = ''
            excel_editor = ''
            excel_publication_date = ''
            excel_date_received = ''
            excel_date_sent_for_update = ''

            # filter only emails, see https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.olobjectclass?view=outlook-pia
            # filter emails that are unread and falls within the date range config
            if (outlook_message.Class == 43) and (outlook_message.UnRead) and \
                (outlook_message.ReceivedTime >= outlook_start_date) and \
                (outlook_message.ReceivedTime <= outlook_end_date):

                try:
                    excel_workbook = openpyxl.load_workbook(currency_tracker_path)
                except:
                    print('Unable to open target currency tracker file')
                else:
                    print('Workbook loaded: ' + str(currency_tracker_path))
                    excel_worksheet_monthly_data = excel_workbook['Monthly Data']
                    excel_worksheet_dashboard = excel_workbook['Dashboard']

                outlook_message_body = outlook_message.Body
                outlook_message_body_parts = outlook_message_body.replace('\r', '')
                outlook_message_body_parts = outlook_message_body_parts.split('\n')

                print('Processing email: ' + outlook_message.Subject)
                excel_publication_date = ''
                excel_date_received = ''
                excel_date_sent_for_update = ''
                excel_deadline = ''
                excel_lag_day = ''
                excel_actual_turnaround_time = ''
                excel_pass_or_fail = ''

                date_sent_for_update = dateparser.parse( str(outlook_message.SentOn) )                
                excel_date_sent_for_update = date_sent_for_update.strftime(excel_date_format)

                for message_body_part in outlook_message_body_parts:
                    if (message_body_part != '') and (':' in message_body_part):
                        (key, value) = message_body_part.split(':')
                        if value != '':
                            if key.lower() == 'user':
                                excel_editor = value.strip()
                            
                            if key.lower() == 'jurisdiction':
                                excel_jurisdiction = value.strip()

                                # standardized the french fedral jurisdiction
                                if excel_jurisdiction.lower() == 'french federal':
                                    excel_jurisdiction = 'federal french'

                                # if the parsed jurisdiction is PEI, expand that into full name
                                if excel_jurisdiction.lower() == 'pei':
                                    excel_jurisdiction = 'prince edward island'

                            if key.lower() == 'type':
                                excel_type = value.strip()

                            print(key + ': ' + value)

                            if 'advance currency' in key.lower():
                                # we parse and format dates if either is in french or english format
                                publication_date = dateparser.parse(value.strip()).date()
                                excel_publication_date = publication_date.strftime(excel_date_format)
                                print("Publication Date: " + excel_publication_date)

                                # check if publication date for specific jurisdiction falls on date 
                                # received schedule found on Link to Lexis.xlsx
                                if publication_date in date_received_schedule[excel_jurisdiction.lower()]:
                                    # if the publication date is same day with date received schedule date, we give 1 day allowance
                                    # this also is applied if date received falls on Sunday                                    
                                    date_received = publication_date
                                else:
                                    # create a special case for nunavut as per mam Rose advise
                                    date_received_sched_override_list = ['nova scotia', 'nunavut', 'northwest territories']
                                    if (excel_jurisdiction.lower() in date_received_sched_override_list):
                                        date_received = publication_date
                                    else:
                                        # need to iterate the nearest date received date
                                        date_received = None
                                        for date_schedule in date_received_schedule[excel_jurisdiction.lower()]:
                                            # the first date that is greater than the pub date then it is set as the date received
                                            if publication_date < date_schedule:
                                                date_received = date_schedule
                                                break
                                            else:
                                                # ignore dates that are older than publication date
                                                pass

                                # for some reason we cannot find the date received schedule so we
                                # we assign the publication date
                                if date_received == None:
                                    date_received = publication_date

                                # if the date received fall on Friday, we count it to receive on Monday or plus 3 days
                                if date_received.isoweekday() == 5:
                                    date_received = date_received + datetime.timedelta(days=3)
                                elif date_received.isoweekday() == 6:
                                    # if the date received fall on a Saturday, we count it to receive on Monday
                                    # or plus 2 days from the date it was received.                                    
                                    date_received = date_received + datetime.timedelta(days=2)
                                else:
                                    # if date received falls on Monday to Thursday, then we add 1 day only
                                    date_received = date_received + datetime.timedelta(days=1)

                                # add 1 day on date received if that falls on holiday
                                if date_received in canada_holidays:
                                    date_received = date_received + datetime.timedelta(days=1)

                                excel_date_received = date_received.strftime(excel_date_format)
                                print("Date received: " + excel_date_received)

                                # how many days the document is completed between date received and date sent for update
                                # we exclude weekends on the TAT conputation
                                excluded_days = [6,7]
                                actual_turnaround_time = 0
                                # create date range between date received and date sent for update
                                # then filter the weekends
                                # we begin counting the TAT next day
                                for date_value in generate_date_range(date_received + datetime.timedelta(days=1), date_sent_for_update.date()):
                                    if date_value.isoweekday() not in excluded_days:
                                        actual_turnaround_time = actual_turnaround_time + 1

                                    # substract 1 day on TAT if the date falls on holiday
                                    if date_value in canada_holidays:
                                        actual_turnaround_time = actual_turnaround_time - 1

                                excel_actual_turnaround_time = actual_turnaround_time
                                print("Actual Turnaround Time: " + str(excel_actual_turnaround_time))

                                # deadline is based on current TAT which is 10 days from the received date
                                excel_deadline = add_business_days(date_received, 10)
                                excel_deadline = excel_deadline.strftime(excel_date_format)
                                print("Deadline: " + excel_deadline)

                                # lag is difference from the date received (actual received of the publications)
                                # and publication date (expected received date)
                                excel_lag_day = date_received - publication_date
                                excel_lag_day = excel_lag_day.days

                                # if turn around time requirements exceeds 10 days, then it is considered failed
                                if excel_actual_turnaround_time <= 10:
                                    excel_pass_or_fail = 'PASS'
                                else:
                                    excel_pass_or_fail = 'FAIL'

                # mark the email as Read
                outlook_message.UnRead = False                

                outlook_message_subject = outlook_message.Subject
                lexisadvance_jurisdiction_code = outlook_message_subject.replace('Update Request for ', '')
                lexisadvance_currency_page = lexisadvance_links[lexisadvance_jurisdiction_code]

                # wait for the page to fully load to avoid getting empty currency date
                chrome_webdriver.implicitly_wait(40)
                chrome_webdriver.get(lexisadvance_currency_page)
                # make a longer waiting period for page to load
                html_elem = WebDriverWait(chrome_webdriver, 40).until(EC.presence_of_element_located((By.NAME, 'HideShowLabel')))
                                                                      
                if "Service is temporarily unavailable." in chrome_webdriver.page_source:
                    raise RuntimeError('Webpage not found: ' + lexisadvance_currency_page)
                else:
                    #html_elem = chrome_webdriver.find_element(By.XPATH, "//h2[@class='SS_HideShowSection SS_Expandable']/Span[@name='HideShowLabel']")
                    html_elem = chrome_webdriver.find_element(By.NAME, 'HideShowLabel')
                    if html_elem is not None:
                        html_elem_dashboard_publication_date = html_elem.text
                        html_elem_dashboard_publication_date = html_elem_dashboard_publication_date.replace('Current to ', '')

                        # french date
                        if 'À jour en date du ' in html_elem_dashboard_publication_date:
                            html_elem_dashboard_publication_date = html_elem_dashboard_publication_date.replace('À jour en date du ', '')

                        if 'À jour en date ' in html_elem_dashboard_publication_date:
                            html_elem_dashboard_publication_date = html_elem_dashboard_publication_date.replace('À jour en date ', '')

                        excel_dashboard_publication_date = dateparser.parse(html_elem_dashboard_publication_date.strip())

                        if excel_dashboard_publication_date is not None:
                            excel_dashboard_publication_date = excel_dashboard_publication_date.strftime("%m/%d/%Y")
                            print('Converted LexisNexis currency date: ' + excel_dashboard_publication_date)
                        else:
                            print('Cannot convert date: ' + html_elem_dashboard_publication_date)

                        excel_dashboard_jurisdiction = excel_jurisdiction.lower()

                        # @note we do not yet have any config for SEC and Annuals
                        if not (excel_dashboard_jurisdiction in jurisdiction_list):
                            raise LookupError('Parsed jurisdiction \''+excel_dashboard_jurisdiction+'\' is not on jurisdiction list. Unable to update the dashboard data')
                        else:
                            # @note accessing the same statutes and regulations webpage more than once cause the tool to
                            # get empty data but manually accessing the page shows data
                            if excel_dashboard_publication_date is not None:
                                if excel_dashboard_jurisdiction == 'alberta':
                                    #AB Stats and AB Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I33'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I34'].value = excel_dashboard_publication_date

                                elif excel_dashboard_jurisdiction == 'british columbia':
                                    #BC Stats and BC Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I35'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I36'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'ontario':
                                    #ON Stats and ON Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I37'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I38'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'saskatchewan':
                                    #SK Stats and SK Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I39'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I40'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'nova scotia':
                                    #NS Stats and NS Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I41'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I42'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'nunavut':
                                    #NU Stats and NU Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I43'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I44'].value = excel_dashboard_publication_date
                                elif (excel_dashboard_jurisdiction == 'new brunswick english') or (excel_dashboard_jurisdiction == 'new brunswick french'):
                                    #NB Stats and NB Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I45'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I46'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'manitoba':
                                    #MB Stats and MB Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I48'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I49'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'northwest territories':
                                    #NT Stats and NT Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I50'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I51'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'newfoundland':
                                    #NL Stats and NL Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I52'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I53'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'prince edward island':
                                    #PE Stats and PE Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I54'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I55'].value = excel_dashboard_publication_date
                                elif excel_dashboard_jurisdiction == 'yukon':
                                    #YT Stats and YT Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I56'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I57'].value = excel_dashboard_publication_date
                                elif (excel_dashboard_jurisdiction == 'federal english') or (excel_dashboard_jurisdiction == 'federal french'):
                                    #FED Stats and FED Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I58'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I59'].value = excel_dashboard_publication_date
                                elif (excel_dashboard_jurisdiction == 'québec english') or (excel_dashboard_jurisdiction == 'québec french'):
                                    #QC Stats and QC Regs
                                    if excel_type == 'Statutes':
                                        excel_worksheet_dashboard['I60'].value = excel_dashboard_publication_date
                                    elif excel_type == 'Regulations':
                                        excel_worksheet_dashboard['I61'].value = excel_dashboard_publication_date

                if excel_type == 'Statutes':
                    # add entry for STAT Amendments
                    excel_insertion_row = str(excel_row_matrix[outlook_target_month.lower()][excel_jurisdiction.lower()]['STAT Amendments'])
                    publication_date_target_column = 'J'
                    date_received_target_column = 'K'
                    date_sent_for_update_target_column = 'L'
                    pass_or_fail_target_column = 'M'
                    lag_target_column = 'N'
                    deadline_target_column = 'O'
                    actual_turnaround_time_target_column = 'P'

                    # we check if entry 1 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'Q'
                        date_received_target_column = 'R'
                        date_sent_for_update_target_column = 'S'
                        pass_or_fail_target_column = 'T'
                        lag_target_column = 'U'
                        deadline_target_column = 'V'
                        actual_turnaround_time_target_column = 'W'
                    
                    # we check if entry 2 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'X'
                        date_received_target_column = 'Y'
                        date_sent_for_update_target_column = 'Z'
                        pass_or_fail_target_column = 'AA'
                        lag_target_column = 'AB'
                        deadline_target_column = 'AC'
                        actual_turnaround_time_target_column = 'AD'
                    
                    # we check if entry 3 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AE'
                        date_received_target_column = 'AF'
                        date_sent_for_update_target_column = 'AG'
                        pass_or_fail_target_column = 'AH'
                        lag_target_column = 'AI'
                        deadline_target_column = 'AJ'
                        actual_turnaround_time_target_column = 'AK'

                    # we check if entry 4 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AL'
                        date_received_target_column = 'AM'
                        date_sent_for_update_target_column = 'AN'
                        pass_or_fail_target_column = 'AO'
                        lag_target_column = 'AP'
                        deadline_target_column = 'AQ'
                        actual_turnaround_time_target_column = 'AR'

                    # we check if entry 5 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AS'
                        date_received_target_column = 'AT'
                        date_sent_for_update_target_column = 'AU'
                        pass_or_fail_target_column = 'AV'
                        lag_target_column = 'AW'
                        deadline_target_column = 'AX'
                        actual_turnaround_time_target_column = 'AY'

                    # we check if entry 6 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AZ'
                        date_received_target_column = 'BA'
                        date_sent_for_update_target_column = 'BB'
                        pass_or_fail_target_column = 'BC'
                        lag_target_column = 'BD'
                        deadline_target_column = 'BE'
                        actual_turnaround_time_target_column = 'BF'

                    # we check if entry 7 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BG'
                        date_received_target_column = 'BH'
                        date_sent_for_update_target_column = 'BI'
                        pass_or_fail_target_column = 'BJ'
                        lag_target_column = 'BK'
                        deadline_target_column = 'BM'
                        actual_turnaround_time_target_column = 'BM'

                    # we check if entry 8 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BN'
                        date_received_target_column = 'BO'
                        date_sent_for_update_target_column = 'BP'
                        pass_or_fail_target_column = 'BQ'
                        lag_target_column = 'BR'
                        deadline_target_column = 'BS'
                        actual_turnaround_time_target_column = 'BT'

                    # we check if entry 9 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BU'
                        date_received_target_column = 'BV'
                        date_sent_for_update_target_column = 'BW'
                        pass_or_fail_target_column = 'BX'
                        lag_target_column = 'BY'
                        deadline_target_column = 'BZ'
                        actual_turnaround_time_target_column = 'CA'

                    excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value = excel_publication_date
                    excel_worksheet_monthly_data[date_received_target_column + excel_insertion_row].value = excel_date_received
                    excel_worksheet_monthly_data[date_sent_for_update_target_column + excel_insertion_row].value = excel_date_sent_for_update
                    excel_worksheet_monthly_data[pass_or_fail_target_column + excel_insertion_row].value = excel_pass_or_fail
                    excel_worksheet_monthly_data[lag_target_column + excel_insertion_row].value = excel_lag_day
                    excel_worksheet_monthly_data[deadline_target_column + excel_insertion_row].value = excel_deadline
                    excel_worksheet_monthly_data[actual_turnaround_time_target_column + excel_insertion_row].value = excel_actual_turnaround_time

                    # add entry for STAT New Doc
                    excel_insertion_row = str(excel_row_matrix[outlook_target_month.lower()][excel_jurisdiction.lower()]['STAT New Doc'])
                    excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value = excel_publication_date
                    excel_worksheet_monthly_data[date_received_target_column + excel_insertion_row].value = excel_date_received
                    excel_worksheet_monthly_data[date_sent_for_update_target_column + excel_insertion_row].value = excel_date_sent_for_update
                    excel_worksheet_monthly_data[pass_or_fail_target_column + excel_insertion_row].value = excel_pass_or_fail
                    excel_worksheet_monthly_data[lag_target_column + excel_insertion_row].value = excel_lag_day
                    excel_worksheet_monthly_data[deadline_target_column + excel_insertion_row].value = excel_deadline
                    excel_worksheet_monthly_data[actual_turnaround_time_target_column + excel_insertion_row].value = excel_actual_turnaround_time
                    
                elif excel_type == 'Regulations':
                    # add entry for REG Amendments
                    excel_insertion_row = str(excel_row_matrix[outlook_target_month.lower()][excel_jurisdiction.lower()]['REG Amendments'])
                    publication_date_target_column = 'J'
                    date_received_target_column = 'K'
                    date_sent_for_update_target_column = 'L'
                    pass_or_fail_target_column = 'M'
                    lag_target_column = 'N'
                    deadline_target_column = 'O'
                    actual_turnaround_time_target_column = 'P'
                                        
                    # we check if entry 1 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'Q'
                        date_received_target_column = 'R'
                        date_sent_for_update_target_column = 'S'
                        pass_or_fail_target_column = 'T'
                        lag_target_column = 'U'
                        deadline_target_column = 'V'
                        actual_turnaround_time_target_column = 'W'
                    
                    # we check if entry 2 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'X'
                        date_received_target_column = 'Y'
                        date_sent_for_update_target_column = 'Z'
                        pass_or_fail_target_column = 'AA'
                        lag_target_column = 'AB'
                        deadline_target_column = 'AC'
                        actual_turnaround_time_target_column = 'AD'
                    
                    # we check if entry 3 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AE'
                        date_received_target_column = 'AF'
                        date_sent_for_update_target_column = 'AG'
                        pass_or_fail_target_column = 'AH'
                        lag_target_column = 'AI'
                        deadline_target_column = 'AJ'
                        actual_turnaround_time_target_column = 'AK'

                    # we check if entry 4 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AL'
                        date_received_target_column = 'AM'
                        date_sent_for_update_target_column = 'AN'
                        pass_or_fail_target_column = 'AO'
                        lag_target_column = 'AP'
                        deadline_target_column = 'AQ'
                        actual_turnaround_time_target_column = 'AR'

                    # we check if entry 5 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AS'
                        date_received_target_column = 'AT'
                        date_sent_for_update_target_column = 'AU'
                        pass_or_fail_target_column = 'AV'
                        lag_target_column = 'AW'
                        deadline_target_column = 'AX'
                        actual_turnaround_time_target_column = 'AY'

                    # we check if entry 6 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AZ'
                        date_received_target_column = 'BA'
                        date_sent_for_update_target_column = 'BB'
                        pass_or_fail_target_column = 'BC'
                        lag_target_column = 'BD'
                        deadline_target_column = 'BE'
                        actual_turnaround_time_target_column = 'BF'

                    # we check if entry 7 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BG'
                        date_received_target_column = 'BH'
                        date_sent_for_update_target_column = 'BI'
                        pass_or_fail_target_column = 'BJ'
                        lag_target_column = 'BK'
                        deadline_target_column = 'BM'
                        actual_turnaround_time_target_column = 'BM'

                    # we check if entry 8 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BN'
                        date_received_target_column = 'BO'
                        date_sent_for_update_target_column = 'BP'
                        pass_or_fail_target_column = 'BQ'
                        lag_target_column = 'BR'
                        deadline_target_column = 'BS'
                        actual_turnaround_time_target_column = 'BT'

                    # we check if entry 9 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BU'
                        date_received_target_column = 'BV'
                        date_sent_for_update_target_column = 'BW'
                        pass_or_fail_target_column = 'BX'
                        lag_target_column = 'BY'
                        deadline_target_column = 'BZ'
                        actual_turnaround_time_target_column = 'CA'

                    excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value = excel_publication_date
                    excel_worksheet_monthly_data[date_received_target_column + excel_insertion_row].value = excel_date_received
                    excel_worksheet_monthly_data[date_sent_for_update_target_column + excel_insertion_row].value = excel_date_sent_for_update
                    excel_worksheet_monthly_data[pass_or_fail_target_column + excel_insertion_row].value = excel_pass_or_fail
                    excel_worksheet_monthly_data[lag_target_column + excel_insertion_row].value = excel_lag_day
                    excel_worksheet_monthly_data[deadline_target_column + excel_insertion_row].value = excel_deadline
                    excel_worksheet_monthly_data[actual_turnaround_time_target_column + excel_insertion_row].value = excel_actual_turnaround_time

                    # add entry for REG New Doc
                    excel_insertion_row = str(excel_row_matrix[outlook_target_month.lower()][excel_jurisdiction.lower()]['REG New Doc'])
                    excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value = excel_publication_date
                    excel_worksheet_monthly_data[date_received_target_column + excel_insertion_row].value = excel_date_received
                    excel_worksheet_monthly_data[date_sent_for_update_target_column + excel_insertion_row].value = excel_date_sent_for_update
                    excel_worksheet_monthly_data[pass_or_fail_target_column + excel_insertion_row].value = excel_pass_or_fail
                    excel_worksheet_monthly_data[lag_target_column + excel_insertion_row].value = excel_lag_day
                    excel_worksheet_monthly_data[deadline_target_column + excel_insertion_row].value = excel_deadline
                    excel_worksheet_monthly_data[actual_turnaround_time_target_column + excel_insertion_row].value = excel_actual_turnaround_time

                elif excel_type == 'Annuals':
                    excel_insertion_row = str(excel_row_matrix[outlook_target_month.lower()][excel_jurisdiction.lower()]['Annuals'])
                    publication_date_target_column = 'J'
                    date_received_target_column = 'K'
                    date_sent_for_update_target_column = 'L'
                    pass_or_fail_target_column = 'M'
                    lag_target_column = 'N'
                    deadline_target_column = 'O'
                    actual_turnaround_time_target_column = 'P'

                    # we check if entry 1 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'Q'
                        date_received_target_column = 'R'
                        date_sent_for_update_target_column = 'S'
                        pass_or_fail_target_column = 'T'
                        lag_target_column = 'U'
                        deadline_target_column = 'V'
                        actual_turnaround_time_target_column = 'W'
                    
                    # we check if entry 2 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'X'
                        date_received_target_column = 'Y'
                        date_sent_for_update_target_column = 'Z'
                        pass_or_fail_target_column = 'AA'
                        lag_target_column = 'AB'
                        deadline_target_column = 'AC'
                        actual_turnaround_time_target_column = 'AD'
                    
                    # we check if entry 3 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AE'
                        date_received_target_column = 'AF'
                        date_sent_for_update_target_column = 'AG'
                        pass_or_fail_target_column = 'AH'
                        lag_target_column = 'AI'
                        deadline_target_column = 'AJ'
                        actual_turnaround_time_target_column = 'AK'

                    # we check if entry 4 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AL'
                        date_received_target_column = 'AM'
                        date_sent_for_update_target_column = 'AN'
                        pass_or_fail_target_column = 'AO'
                        lag_target_column = 'AP'
                        deadline_target_column = 'AQ'
                        actual_turnaround_time_target_column = 'AR'

                    # we check if entry 5 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AS'
                        date_received_target_column = 'AT'
                        date_sent_for_update_target_column = 'AU'
                        pass_or_fail_target_column = 'AV'
                        lag_target_column = 'AW'
                        deadline_target_column = 'AX'
                        actual_turnaround_time_target_column = 'AY'

                    # we check if entry 6 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AZ'
                        date_received_target_column = 'BA'
                        date_sent_for_update_target_column = 'BB'
                        pass_or_fail_target_column = 'BC'
                        lag_target_column = 'BD'
                        deadline_target_column = 'BE'
                        actual_turnaround_time_target_column = 'BF'

                    # we check if entry 7 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BG'
                        date_received_target_column = 'BH'
                        date_sent_for_update_target_column = 'BI'
                        pass_or_fail_target_column = 'BJ'
                        lag_target_column = 'BK'
                        deadline_target_column = 'BM'
                        actual_turnaround_time_target_column = 'BM'

                    # we check if entry 8 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BN'
                        date_received_target_column = 'BO'
                        date_sent_for_update_target_column = 'BP'
                        pass_or_fail_target_column = 'BQ'
                        lag_target_column = 'BR'
                        deadline_target_column = 'BS'
                        actual_turnaround_time_target_column = 'BT'

                    # we check if entry 9 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BU'
                        date_received_target_column = 'BV'
                        date_sent_for_update_target_column = 'BW'
                        pass_or_fail_target_column = 'BX'
                        lag_target_column = 'BY'
                        deadline_target_column = 'BZ'
                        actual_turnaround_time_target_column = 'CA'

                    excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value = excel_publication_date
                    excel_worksheet_monthly_data[date_received_target_column + excel_insertion_row].value = excel_date_received
                    excel_worksheet_monthly_data[date_sent_for_update_target_column + excel_insertion_row].value = excel_date_sent_for_update
                    excel_worksheet_monthly_data[pass_or_fail_target_column + excel_insertion_row].value = excel_pass_or_fail
                    excel_worksheet_monthly_data[lag_target_column + excel_insertion_row].value = excel_lag_day
                    excel_worksheet_monthly_data[deadline_target_column + excel_insertion_row].value = excel_deadline
                    excel_worksheet_monthly_data[actual_turnaround_time_target_column + excel_insertion_row].value = excel_actual_turnaround_time

                elif excel_type == 'SEC':
                    excel_insertion_row = str(excel_row_matrix[outlook_target_month.lower()][excel_jurisdiction.lower()]['SEC'])
                    publication_date_target_column = 'J'
                    date_received_target_column = 'K'
                    date_sent_for_update_target_column = 'L'
                    pass_or_fail_target_column = 'M'
                    lag_target_column = 'N'
                    deadline_target_column = 'O'
                    actual_turnaround_time_target_column = 'P'

                    # we check if entry 1 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'Q'
                        date_received_target_column = 'R'
                        date_sent_for_update_target_column = 'S'
                        pass_or_fail_target_column = 'T'
                        lag_target_column = 'U'
                        deadline_target_column = 'V'
                        actual_turnaround_time_target_column = 'W'
                    
                    # we check if entry 2 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'X'
                        date_received_target_column = 'Y'
                        date_sent_for_update_target_column = 'Z'
                        pass_or_fail_target_column = 'AA'
                        lag_target_column = 'AB'
                        deadline_target_column = 'AC'
                        actual_turnaround_time_target_column = 'AD'
                    
                    # we check if entry 3 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AE'
                        date_received_target_column = 'AF'
                        date_sent_for_update_target_column = 'AG'
                        pass_or_fail_target_column = 'AH'
                        lag_target_column = 'AI'
                        deadline_target_column = 'AJ'
                        actual_turnaround_time_target_column = 'AK'

                    # we check if entry 4 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AL'
                        date_received_target_column = 'AM'
                        date_sent_for_update_target_column = 'AN'
                        pass_or_fail_target_column = 'AO'
                        lag_target_column = 'AP'
                        deadline_target_column = 'AQ'
                        actual_turnaround_time_target_column = 'AR'

                    # we check if entry 5 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AS'
                        date_received_target_column = 'AT'
                        date_sent_for_update_target_column = 'AU'
                        pass_or_fail_target_column = 'AV'
                        lag_target_column = 'AW'
                        deadline_target_column = 'AX'
                        actual_turnaround_time_target_column = 'AY'

                    # we check if entry 6 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'AZ'
                        date_received_target_column = 'BA'
                        date_sent_for_update_target_column = 'BB'
                        pass_or_fail_target_column = 'BC'
                        lag_target_column = 'BD'
                        deadline_target_column = 'BE'
                        actual_turnaround_time_target_column = 'BF'

                    # we check if entry 7 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BG'
                        date_received_target_column = 'BH'
                        date_sent_for_update_target_column = 'BI'
                        pass_or_fail_target_column = 'BJ'
                        lag_target_column = 'BK'
                        deadline_target_column = 'BM'
                        actual_turnaround_time_target_column = 'BM'

                    # we check if entry 8 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BN'
                        date_received_target_column = 'BO'
                        date_sent_for_update_target_column = 'BP'
                        pass_or_fail_target_column = 'BQ'
                        lag_target_column = 'BR'
                        deadline_target_column = 'BS'
                        actual_turnaround_time_target_column = 'BT'

                    # we check if entry 9 has contents
                    if excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value is not None:
                        publication_date_target_column = 'BU'
                        date_received_target_column = 'BV'
                        date_sent_for_update_target_column = 'BW'
                        pass_or_fail_target_column = 'BX'
                        lag_target_column = 'BY'
                        deadline_target_column = 'BZ'
                        actual_turnaround_time_target_column = 'CA'

                    excel_worksheet_monthly_data[publication_date_target_column + excel_insertion_row].value = excel_publication_date
                    excel_worksheet_monthly_data[date_received_target_column + excel_insertion_row].value = excel_date_received
                    excel_worksheet_monthly_data[date_sent_for_update_target_column + excel_insertion_row].value = excel_date_sent_for_update
                    excel_worksheet_monthly_data[pass_or_fail_target_column + excel_insertion_row].value = excel_pass_or_fail
                    excel_worksheet_monthly_data[lag_target_column + excel_insertion_row].value = excel_lag_day
                    excel_worksheet_monthly_data[deadline_target_column + excel_insertion_row].value = excel_deadline
                    excel_worksheet_monthly_data[actual_turnaround_time_target_column + excel_insertion_row].value = excel_actual_turnaround_time                    
                
                excel_workbook.save(currency_tracker_path)
                print('Workbook saved: ' + str(currency_tracker_path))
                print("-----------------------\n")

# clean up Chrome webdriver session
chrome_webdriver.close()

# wait 5 seconds
for x in range(6):
    pass
print('Finished processing files. Please close this window.')
