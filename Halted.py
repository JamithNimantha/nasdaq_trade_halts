# import libraries

from datetime import datetime

try:
    from selenium import webdriver
    import time

    # Firefox options
    from selenium.webdriver.firefox.options import Options
    # Un-Comment this when using the chromedriver
    # from selenium.webdriver.chrome.options import Options
    from selenium.common.exceptions import NoSuchElementException
    import psycopg2
    import csv
    import datetime as dt
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from datetime import date
    import xlrd

except Exception as e:
    print(e)


def getData():
    """
    Function to get the table from the website

    Returns:
        [dictionary] : A dictionary containing data from the the website
    """

    # Path for chromedriver
    # path = 'C:\Windows\chromedriver.exe'

    # Path for geckodriver
    path = "C:\Windows\geckodriver.exe"
    url = 'https://nasdaqtrader.com/Trader.aspx?id=TradeHalts'

    # Chromedriver settings
    # option = Options()
    # option.add_argument("--disable-gpu")
    # option.add_argument("--headless")
    # driver = webdriver.Chrome(executable_path = path, options= option)

    # Firefox settings
    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options, executable_path=path)

    driver.get(url)

    # Data to be stored in a dictionary
    data = {}

    # Table Rows
    tr = driver.find_elements_by_tag_name('tr')
    # Table headings
    th = driver.find_elements_by_tag_name('th')

    for heading in th:
        data[heading.text] = []

    for tc in tr[3:-1]:
        for count, x in enumerate(tc.find_elements_by_tag_name('td')):
            data[th[count].text].append(x.text)

    driver.quit()

    return data


def database():
    """Function to connect to the PostgreSQL Database

    Returns:
        [object]: Cursor object
    """

    # Read Control.csv
    csv_data = dict(csv.reader(open('Control\Control.csv')))

    # Assign values from control.csv
    database = csv_data["Database"]
    user = csv_data["User Name"]
    pw = csv_data["Password"]
    host = csv_data["Database Host"]

    # Connecting to DB and Cursor creation
    conn = psycopg2.connect(database=database, user=user, password=pw, host=host, port="5432")
    conn.autocommit = True

    return conn.cursor()


def ReadExcel(symbol):
    """Function to get NewsExcel data for a symbol

    Args:
        symbol (string): Symbol for a particular row

    Returns:
        [Tuple]: Returns Volume, Float short, days, M cap, Price
    """

    # Read control.csv
    control = dict(csv.reader(open('Control\Control.csv')))

    # Getting column numbers from control.csv
    vol, flt_srt, days, mcap, price = (int(control['News_Excel_File.xlsx Column Number for Volume in Thousands']) - 1,
                                       int(control['News_Excel_File.xlsx Column Number for Float Short']) - 1,
                                       int(control['News_Excel_File.xlsx Column Number for Days']) - 1,
                                       int(control[
                                               'News_Excel_File.xlsx Column Number for Market Cap in Millions']) - 1,
                                       int(control['News_Excel_File.xlsx Column Number for Price']) - 1)
    content = []

    # Loading the excel file
    wb_load = xlrd.open_workbook(f"{control['News_Excel_File.xlsx Location']}")
    wb = wb_load.sheet_by_index(0)

    for ro in range(0, wb.nrows):
        key1 = wb.cell(ro, 0)

        # If symbol is found tuple is returned with values from News Excel
        if symbol == key1.value:
            content = (int(wb.cell(ro, vol).value), wb.cell(ro, flt_srt).value, int(wb.cell(ro, days).value),
                       int(wb.cell(ro, mcap).value), int(wb.cell(ro, price).value))
            break

        # Else, tuple is returned with NA values
        else:
            content = ('NA', 'NA', 'NA', 'NA', 'NA')

    return content


def ReadExcelMore(symbol):
    """Function to get NewsExcel data for a symbol

    Args:
        symbol (string): Symbol for a particular row

    Returns:
        [Tuple]: Returns Volume, Float short, days, M cap, Price
    """

    # Read control.csv
    control = dict(csv.reader(open('Control\Control.csv')))

    # Getting column numbers from control.csv
    label, ind, opt, cashburn = (int(control['News_Excel_File.xlsx Column Number for Label']) - 1,
                                 int(control['News_Excel_File.xlsx Column Number for Industry']) - 1,
                                 int(control['News_Excel_File.xlsx Column Number for OPT']) - 1,
                                 int(control['News_Excel_File.xlsx Column Number for Cash Burn Mnth']) - 1)

    # Loading the excel file
    wb_load = xlrd.open_workbook(f"{control['News_Excel_File.xlsx Location']}")
    wb = wb_load.sheet_by_index(0)

    for ro in range(0, wb.nrows):
        key1 = wb.cell(ro, 0)

        # If symbol is found tuple is returned with values from News Excel
        if symbol == key1.value:
            content = (
            wb.cell(ro, label).value, wb.cell(ro, ind).value, wb.cell(ro, opt).value, int(wb.cell(ro, cashburn).value))
            break

        # Else, tuple is returned with NA values
        else:
            content = ('NA', 'NA', 'NA', 'NA')

    return content


def updateExcel(data, vol, flt_sht, days, price, mcap):
    """Function to update the CSV file 

    Args:
        data (dictionary): Dictionary with row data
        vol (int): Volume from News Excel
        flt_sht (Float): Float short from news excel
        days (int): Days from news excel
        price (int): price from news excel
        mcap (int): M cap from news excel
    """

    control = dict(csv.reader(open('Control\Control.csv')))
    content = list(csv.reader(open(f'{control["Halts.csv Location"]}')))

    to_edit = [str(data['Halt Date']), data['Halt Time'],
               data['Issue Symbol'], data['Market'], data['Reason Codes'].replace('\n', ''),
               data['Resumption Date'], data['Resumption Quote Time'], data['Resumption Trade Time'], vol, flt_sht,
               days, price, mcap]
    to_add = False

    for count, row in enumerate(content):

        # If symbol data is found, the row is edited
        if data['Issue Symbol'] in row:
            content[count] = to_edit
            to_add = False
            break
        else:
            to_add = True

    if to_add:
        content.append(to_edit)

    f = open(f'{control["Halts.csv Location"]}', mode='w', newline='')
    wr = csv.writer(f)
    wr.writerows(content)
    f.close()


def checks(data):
    cursor = database()
    control = dict(csv.reader(open('Control\Control.csv')))

    # Iterate through all the rows
    for count, symbol in enumerate(data['Issue Symbol']):

        # Get today's date
        today = date.today()

        # Get excel data
        vol, flt_srt, days, mcap, price = ReadExcel(symbol)

        HaltDate, HaltTime, Symbol, Market, IssueName, ReasonCode, ThresholdPrice, ResumptionDate, ResumptionQuoteTime, ResumptionTradeTime = (
        data['Halt Date'][count], data['Halt Time'][count],
        data['Issue Symbol'][count], data['Market'][count], data['Issue Name'], data['Reason Codes'][count],
        data['Pause Threshold Price'][count],
        data['Resumption Date'][count], data['Resumption Quote Time'][count], data['Resumption Trade Time'][count])

        # Halt Date for the row
        HaltDate = dt.datetime.strptime(HaltDate, "%m/%d/%Y").date()

        # Dictionary that will be updated in mails and csv
        to_send = {'Halt Date': HaltDate, 'Halt Time': HaltTime, 'Issue Symbol': symbol, 'Market': Market,
                   'Issue Name': IssueName, 'Reason Codes': ReasonCode, 'Pause Threshold Price': ThresholdPrice,
                   'Resumption Date': ResumptionDate, 'Resumption Quote Time': ResumptionQuoteTime,
                   'Resumption Trade Time': ResumptionTradeTime}

        # Check 1: If row's Halt Date < Today's date, We will move on to the next row.
        if today < HaltDate:
            print('Checks not passed for', symbol)
            continue

        # Get data for the row from database using the primary keys. If row doesn't exist and empty list is returned.
        cursor.execute("SELECT * FROM public.halts WHERE halt_date = %s AND halt_time = %s AND symbol = %s",
                       (str(HaltDate), HaltTime, symbol))
        symbol_data = cursor.fetchall()

        # Check 2: If the row is not found in database.
        if symbol_data == []:

            # Check : If Res Quote time or date is specified
            if ResumptionQuoteTime != '' or ResumptionDate != '':

                try:

                    # Check : If volume specified in News Excel is greated than volume specified in News Excel.
                    if vol > int(control['Minimum Volume in Thousands']):
                        SendMail(to_send)
                    else:
                        print('Checks not passed for', symbol)

                # If vol is NA
                except TypeError:
                    SendMail(to_send)

                # Convert to datetime object
                ResumptionDate = dt.datetime.strptime(ResumptionDate, "%m/%d/%Y").date()

            # Check: if Res quote time or date not specified
            else:
                ResumptionDate = None
                ResumptionQuoteTime = None
                print('Checks not passed for', symbol)

            if ResumptionTradeTime == '':
                ResumptionTradeTime = None

            # Insert the row in database
            cursor.execute(
                "INSERT INTO public.halts(halt_date, halt_time, symbol, market, code, resume_date, resume_quote_time, resume_trade_time) VALUES(%s, %s, %s, %s, %s, %s, %s, %s)",
                (HaltDate, HaltTime, Symbol, Market, ReasonCode, ResumptionDate, ResumptionQuoteTime,
                 ResumptionTradeTime))

        # Check 3: If the row is found in database
        else:

            # Symbol data[0][6] is Resumption quote time and Symbol data[0][7] is resumption trade time from database
            # Check: If res quote time or res trade time specified in database
            if symbol_data[0][6] != None or symbol_data[0][7] != None:

                # If Resumption Quote time or resumption trade time are different in extracted data
                if str(symbol_data[0][6]) != ResumptionQuoteTime or str(symbol_data[0][7]) != ResumptionTradeTime:
                    SendMail(to_send)

                else:
                    print('Checks not passed for', symbol)

                if ResumptionTradeTime == '':
                    ResumptionTradeTime = None

            # Check: If resumption quote time or res trade time not specified in databse
            else:

                # Check: If resumption quote time or resumption trade time specified now
                if ResumptionQuoteTime != '' or ResumptionTradeTime != '':
                    ResumptionDate = dt.datetime.strptime(ResumptionDate, "%m/%d/%Y").date()
                    SendMail(to_send)

                    if ResumptionTradeTime == '':
                        ResumptionTradeTime = None

                else:
                    ResumptionQuoteTime = None
                    ResumptionDate = None
                    ResumptionTradeTime = None
                    print('Checks not passed for', symbol)

            # Update row
            cursor.execute(
                "UPDATE public.halts SET resume_date = %s, resume_quote_time = %s, resume_trade_time = %s WHERE halt_date = %s AND halt_time = %s AND symbol = %s",
                (ResumptionDate, ResumptionQuoteTime, ResumptionTradeTime, str(HaltDate), HaltTime, symbol))

        try:
            flt_srt = float('%.2f' % float(flt_srt))
        except ValueError:
            flt_srt = 'NA'

        # Update the CSV
        updateExcel(to_send, vol, flt_srt, days, price, mcap)


def SendMail(data):
    # Read control.csv
    today = date.today()
    yesterday = today - dt.timedelta(days=1)
    cursor = database()
    cursor.execute("""SELECT timestamp,headline, url FROM public.news_headlines
WHERE ( entry_date = %s OR entry_date = %s ) AND 
( symbol_1 = %s
OR symbol_1 = %s
OR symbol_2 = %s
OR symbol_3 = %s
OR symbol_4 = %s
OR symbol_5 = %s
OR symbol_6 = %s
OR symbol_7 = %s
OR symbol_8 = %s
OR symbol_9 = %s
OR symbol_10 = %s)
ORDER BY timestamp DESC
""", (today.strftime("%Y-%m-%d"), yesterday.strftime("%Y-%m-%d"), data['Issue Symbol'], data['Issue Symbol'],
      data['Issue Symbol'], data['Issue Symbol'], data['Issue Symbol'], data['Issue Symbol'], data['Issue Symbol'],
      data['Issue Symbol'], data['Issue Symbol'], data['Issue Symbol'], data['Issue Symbol']))
    news = cursor.fetchall()

    if news == []:
        timestamp, headline, url = 'NA', 'NA', 'NA'
    else:
        timestamp, headline, url = news[0][0], news[0][1], news[0][2]

    control = dict(csv.reader(open('Control\Control.csv')))

    msg = MIMEMultipart()
    msg['From'] = control['Email SMTP ID']
    msg['To'] = control['Email TO Email ID']

    # Reason Code
    ReasonCode = data['Reason Codes'].replace("\n", "")

    vol, flt_srt, days, mcap, price = ReadExcel(data['Issue Symbol'])

    # If data is found in excel for  the symbol
    if vol != 'NA':

        label, ind, opt, cashburn = ReadExcelMore(data['Issue Symbol'])

        try:
            flt_srt = float('%.2f' % float(flt_srt))
        except ValueError:
            flt_srt = 0

        subject = f""":Nasdaq Halt: {data['Issue Symbol']} :NEWS; {headline} :Res; {data['Resumption Trade Time']} :Code; {ReasonCode} :Vol; {vol} :Flt Sht; {flt_srt} :Days; {days} :McapM; ${mcap} :Price; ${price})"""

        # Html structure
        html = f"""
        <p>
        {str(timestamp)}    {headline}
        {url}
        </p>
        <table cellspacing="1.5" border="1">
        <tr>"""

        table_content = ['Time', 'Sym', 'Reason Code', 'Res DT', 'Res QT TM', 'Res TRD TM', 'Vol', 'Flt Sht', 'Days',
                         'Price', 'M cap M', 'Label', 'Industry', 'OPT', 'Cash Burn Mnth']
        for key in table_content:
            html += f""" <th>{key}</th>"""

        html += '</tr>  <tr>'

        html += f"""<td>{data['Halt Time']}</td>"""

        html += f"""<td>{data['Issue Symbol']}</td>"""

        if data['Reason Codes'] == 'LUDP':
            html += f"""<td style='background-color: red;'>{ReasonCode}</td>"""
        elif data['Reason Codes'] == 'M':
            html += f"""<td style='background-color: yellow;'>{ReasonCode}</td>"""
        else:
            html += f"""<td>{ReasonCode}</td>"""

        html += f"""<td>{data['Resumption Date']}</td>"""

        html += f"""<td>{data['Resumption Quote Time']}</td>"""

        html += f"""<td>{data['Resumption Trade Time']}</td>"""

        if int(vol) <= 400:
            html += f"""<td style='background-color: red;'>{vol}</td>"""
        elif int(vol) > 400 and int(vol) <= 750:
            html += f"""<td style='background-color: yellow;'>{vol}</td>"""
        else:
            html += f"""<td style='background-color: LawnGreen;'>{vol}</td>"""

        if int(flt_srt) > 5:
            html += f"""<td style='background-color: LawnGreen;'>{flt_srt}</td>"""
        else:
            html += f"""<td>{flt_srt}</td>"""

        if int(days) <= 30:
            html += f"""<td style='background-color: LawnGreen;'>{days}</td>"""
        elif int(days) > 30 and int(days) <= 90:
            html += f"""<td style='background-color: yellow;'>{days}</td>"""
        else:
            html += f"""<td>{days}</td>"""

        if int(price) <= 5:
            html += f"""<td style='background-color: red;'>${price}</td>"""
        else:
            html += f"""<td>${price}</td>"""

        if mcap <= 500:
            html += f"""<td style='background-color: LawnGreen;'>${mcap}</td>"""
        else:
            html += f"""<td>${mcap}</td>"""

        html += f"""<td>{label}</td>
        <td>{ind}</td>"""

        if opt.lower() == 'yes':
            html += f"""<td style='background-color: LawnGreen;'>{opt}</td>"""
        else:
            html += f"""<td>{opt}</td>"""

        if round(cashburn) < 24:
            html += f"""<td style='background-color: red;'>{cashburn}</td>"""
        else:
            html += f"""<td>{cashburn}</td> </tr>  </table>"""


    # If data is not found from excel
    else:
        subject = f""":Nasdaq Halt: {data['Issue Symbol']} :NEWS; {headline} :Res; {data['Resumption Trade Time']} :Code; {ReasonCode} :Vol; NA :Flt Sht; NA :Days; NA :McapM; $NA :Price; $NA"""

        # Html structure
        html = f"""
        <p>
        {str(timestamp)}    {headline}
        {url}
        </p>
        <table cellspacing="1.5" border="1">
        <tr>"""

        table_content = ['Time', 'Sym', 'Reason Code', 'Res DT', 'Res QT TM', 'Res TRD TM', 'Vol', 'Flt Sht', 'Days',
                         'Price', 'M cap M', 'Label', 'Industry', 'OPT', 'Cash Burn Mnth']
        for key in table_content:
            html += f""" <th>{key}</th>"""

        html += '</tr>  <tr>'

        html += f"""<td>{data['Halt Time']}</td>"""

        html += f"""<td>{data['Issue Symbol']}</td>"""

        if data['Reason Codes'] == 'LUDP':
            html += f"""<td style='background-color: red;'>{ReasonCode}</td>"""
        elif data['Reason Codes'] == 'M':
            html += f"""<td style='background-color: yellow;'>{ReasonCode}</td>"""
        else:
            html += f"""<td>{ReasonCode}</td>"""

        html += f"""<td>{data['Resumption Date']}</td>"""

        html += f"""<td>{data['Resumption Quote Time']}</td>"""

        html += f"""<td>{data['Resumption Trade Time']}</td>"""

        html += """
        <td>NA</td>
        <td>NA</td>
        <td>NA</td>
        <td>NA</td>
        <td>NA</td>
        <td>NA</td>
        <td>NA</td>
        <td>NA</td>
        <td>$NA</td> </tr>  </table>"""

    msg['Subject'] = subject
    msg.attach(MIMEText(html, 'html'))
    txt = msg.as_string()

    # Check if SPA is required

    try:
        if control['Require logon using Secure Password Authentication (SPA)'].lower() == 'no':
            server = str(control['Email SMTP Server Name / IP Address']) + ":" + str(control['Email SMTP Server Port'])
            s = smtplib.SMTP(server)
        else:
            s = smtplib.SMTP(control['Email SMTP Server Name / IP Address'], int(control['Email SMTP Server Port']))
            s.starttls()

        s.login(control['Email SMTP ID'], control['Email SMTP Password'])
        s.sendmail(control['Email SMTP ID'], control['Email TO Email ID'], txt)
        s.quit()
        print('Sent mail')

    except smtplib.SMTPException as e:
        print(e)
        print("Email couldn't be sent")


while True:
    control = dict(csv.reader(open('Control\Control.csv')))
    if int(dt.datetime.now().strftime("%H")) == int(
            dt.datetime.strptime(control['Time to Stop'], '%I:%S %p').strftime("%H")):
        break

    try:

        loop_time = int(control['Check Frequency in Minutes']) * 60
        data = getData()

        try:
            checks(data)
        except KeyError:
            pass

        time.sleep(loop_time)

    except KeyboardInterrupt:

        print('Exiting')
        break

    except NoSuchElementException:
        continue
