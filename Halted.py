# import libraries

try:
    import os
    import pause as pause
    import time
    import psycopg2
    import csv
    import datetime as dt
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from datetime import date
    import xlrd
    import feedparser

except Exception as e:
    print(e)


def getData():
    """
    Function to get the table from the website

    Returns:
        [dictionary] : A dictionary containing data from the the website
    """

    url = 'http://www.nasdaqtrader.com/rss.aspx?feed=tradehalts'

    data = {'Halt Date': [], 'Halt Time': [], 'Issue Symbol': [],
            'Issue Name': [], 'Market': [],
            'Reason Codes': [], 'Pause Threshold Price': [], 'Resumption Date': [],
            'Resumption Quote Time': [], 'Resumption Trade Time': []}

    feed = feedparser.parse(url)

    print(f'Found {len(feed.entries)} entries!')

    for entry in feed.entries:
        data['Halt Date'].append(entry.ndaq_haltdate)
        data['Halt Time'].append(entry.ndaq_halttime)
        data['Issue Symbol'].append(entry.ndaq_issuesymbol)
        data['Issue Name'].append(entry.ndaq_issuename)
        data['Market'].append(entry.ndaq_market)
        data['Reason Codes'].append(entry.ndaq_reasoncode)
        data['Pause Threshold Price'].append(entry.ndaq_pausethresholdprice)
        data['Resumption Date'].append(entry.ndaq_resumptiondate)
        data['Resumption Quote Time'].append(entry.ndaq_resumptionquotetime)
        data['Resumption Trade Time'].append(entry.ndaq_resumptiontradetime)

    return data


def database():
    """Function to connect to the PostgreSQL Database

    Returns:
        [object]: Cursor object
    """

    # Read Control.csv
    csv_data = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

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
    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

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
    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

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
                wb.cell(ro, label).value, wb.cell(ro, ind).value, wb.cell(ro, opt).value,
                int(wb.cell(ro, cashburn).value))
            break

        # Else, tuple is returned with NA values
        else:
            content = ('NA', 'NA', 'NA', 'NA')

    return content


def ReadExcelAll(symbol):
    """Function to get ALL NewsExcel data by symbol

    Args:
        symbol (string): Symbol for a particular row

    Returns:
        [Tuple]: Returns ALL
    """

    # Read control.csv
    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

    # Getting column numbers from control.csv
    dir, country, stop_diff, from_n_open, perth_week, perf_month, perf_quart, prod_m, n_q1_est_n_eps, \
    est_n_earnm, n_sales_m, n_q_1_esp, chg_nq1_nesp_4_wk, ins_own, ins_trans, inst_trans, ipo_date, earn_date, trg_prc, \
    zacks_rank, z_rank_chg, n_q2_n_eps, n_f1_n_eps, nf1_n_sales_m, p_e, fwd_p_e, p_g_e, p_f_c_f, n_e_p_s_5_yrs, one_min_vol = \
        (2, 4, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 34, 35, 36, 37, 38,
         39)

    # Loading the excel file
    wb_load = xlrd.open_workbook(f"{control['News_Excel_File.xlsx Location']}")
    wb = wb_load.sheet_by_index(0)

    for ro in range(0, wb.nrows):
        key1 = wb.cell(ro, 0)
        # If symbol is found tuple is returned with values from News Excel
        if symbol == key1.value:
            content = (
                wb.cell(ro, dir).value,
                wb.cell(ro, country).value,
                wb.cell(ro, stop_diff).value,
                wb.cell(ro, from_n_open).value,
                wb.cell(ro, perth_week).value,
                wb.cell(ro, perf_month).value,
                wb.cell(ro, perf_quart).value,
                wb.cell(ro, prod_m).value,
                wb.cell(ro, n_q1_est_n_eps).value,
                wb.cell(ro, est_n_earnm).value,
                wb.cell(ro, n_sales_m).value,
                wb.cell(ro, n_q_1_esp).value,
                wb.cell(ro, chg_nq1_nesp_4_wk).value,
                wb.cell(ro, ins_own).value,
                wb.cell(ro, ins_trans).value,
                wb.cell(ro, inst_trans).value,
                get_date(wb.cell(ro, ipo_date).value, wb_load.datemode),
                get_date(wb.cell(ro, earn_date).value, wb_load.datemode),
                wb.cell(ro, trg_prc).value,
                wb.cell(ro, zacks_rank).value,
                wb.cell(ro, z_rank_chg).value,
                wb.cell(ro, n_q2_n_eps).value,
                wb.cell(ro, n_f1_n_eps).value,
                wb.cell(ro, nf1_n_sales_m).value,
                wb.cell(ro, p_e).value,
                wb.cell(ro, fwd_p_e).value,
                wb.cell(ro, p_g_e).value,
                wb.cell(ro, p_f_c_f).value,
                wb.cell(ro, n_e_p_s_5_yrs).value,
                wb.cell(ro, one_min_vol).value)

            break

        # Else, tuple is returned with NA values
        else:
            content = ('NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA',
                       'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA')

    return content


def get_date(value, datemod):
    try:
        return dt.datetime(*xlrd.xldate_as_tuple(value, datemod)).strftime('%m/%d/%Y')
    except TypeError:
        return value


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

    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))
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
    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

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
                except TypeError as e:
                    # SendMail(to_send) commented as This sends an email if the Vol is NA
                    print(e)

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

    control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))

    msg = MIMEMultipart()
    msg['From'] = control['Email SMTP ID']
    msg['To'] = control['Email TO Email ID']

    # Reason Code
    ReasonCode = data['Reason Codes'].replace("\n", "")

    vol, flt_srt, days, mcap, price = ReadExcel(data['Issue Symbol'])

    # If data is found in excel for  the symbol
    if vol != 'NA':

        label, ind, opt, cashburn = ReadExcelMore(data['Issue Symbol'])

        dir_, country, stop_diff, from_n_open, perth_week, perf_month, perf_quart, prod_m, n_q1_est_n_eps, \
        est_n_earnm, n_sales_m, n_q_1_esp, chg_nq1_nesp_4_wk, ins_own, ins_trans, inst_trans, ipo_date, earn_date, trg_prc, \
        zacks_rank, z_rank_chg, n_q2_n_eps, n_f1_n_eps, nf1_n_sales_m, p_e, fwd_p_e, p_g_e, p_f_c_f, n_e_p_s_5_yrs, one_min_vol = \
            ReadExcelAll(data['Issue Symbol'])

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

        table_content = ['Time', 'Sym', 'Reason Code', 'Res DT', 'Res QT TM', 'Res TRD TM',
                         'Dir',
                         'Country',
                         'Stop Diff',
                         'from nOpen',
                         'Perf Week',
                         'Perf Month',
                         'Perf Quart',
                         'Prod M',
                         'nQ1 Est nEPS',
                         'nQ1 Est nEarn M',
                         'nQ1 nSales M',
                         'nQ1 ESP',
                         'Chg nQ1 nEPS 4 Wk',
                         'Ins Own',
                         'Ins Trans',
                         'Inst Trans',
                         'IPO Date',
                         'Earn Date',
                         'Trg Prc %',
                         'Zacks Rank',
                         'ZRank Chg',
                         'nQ2 nEPS',
                         'nF1 nEPS',
                         'nF1 nSales M',
                         'P/E',
                         'Fwd P/E',
                         'PEG',
                         'P/ FCF',
                         'nEPS 5 Yr',
                         'Min Vol', 'Vol', 'Flt Sht', 'Days',
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

        html += f"""<td>{dir_}</td>"""
        html += f"""<td>{country}</td>"""
        try:
            html += f"""<td>${"{:.2f}".format(stop_diff)}</td>"""
        except ValueError:
            html += f"""<td>${stop_diff}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(from_n_open)}</td>"""
        except ValueError:
            html += f"""<td>{from_n_open}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(perth_week)}</td>"""
        except ValueError:
            html += f"""<td>{perth_week}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(perf_month)}</td>"""
        except ValueError:
            html += f"""<td>{perf_month}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(perf_quart)}</td>"""
        except ValueError:
            html += f"""<td>{perf_quart}</td>"""

        try:
            html += f"""<td>${"{:.2f}".format(prod_m)}</td>"""
        except ValueError:
            html += f"""<td>${prod_m}</td>"""

        try:
            html += f"""<td>${"{:.2f}".format(n_q1_est_n_eps)}</td>"""
        except ValueError:
            html += f"""<td>${n_q1_est_n_eps}</td>"""

        try:
            html += f"""<td>${"{:.2f}".format(est_n_earnm)}</td>"""
        except ValueError:
            html += f"""<td>${est_n_earnm}</td>"""

        try:
            html += f"""<td>${"{:.2f}".format(n_sales_m)}</td>"""
        except ValueError:
            html += f"""<td>${n_sales_m}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(n_q_1_esp)}</td>"""
        except ValueError:
            html += f"""<td>{n_q_1_esp}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(chg_nq1_nesp_4_wk)}</td>"""
        except ValueError:
            html += f"""<td>{chg_nq1_nesp_4_wk}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(ins_own)}</td>"""
        except ValueError:
            html += f"""<td>{ins_own}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(ins_trans)}</td>"""
        except ValueError:
            html += f"""<td>{ins_trans}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(inst_trans)}</td>"""
        except ValueError:
            html += f"""<td>{inst_trans}</td>"""

        html += f"""<td>{ipo_date}</td>"""
        html += f"""<td>{earn_date}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(trg_prc)}</td>"""
        except ValueError:
            html += f"""<td>{trg_prc}</td>"""

        html += f"""<td>{zacks_rank}</td>"""
        html += f"""<td>{z_rank_chg}</td>"""
        html += f"""<td>{n_q2_n_eps}</td>"""

        try:
            html += f"""<td>${"{:.2f}".format(n_f1_n_eps)}</td>"""
        except ValueError:
            html += f"""<td>${n_f1_n_eps}</td>"""

        try:
            html += f"""<td>${"{:.2f}".format(nf1_n_sales_m)}</td>"""
        except ValueError:
            html += f"""<td>${nf1_n_sales_m}</td>"""

        html += f"""<td>{p_e}</td>"""
        html += f"""<td>{fwd_p_e}</td>"""
        html += f"""<td>{p_g_e}</td>"""
        html += f"""<td>{p_f_c_f}</td>"""

        try:
            html += f"""<td>{"{:.2%}".format(n_e_p_s_5_yrs)}</td>"""
        except ValueError:
            html += f"""<td>{n_e_p_s_5_yrs}</td>"""

        html += f"""<td>{one_min_vol}</td>"""

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


control = dict(csv.reader(open(f'Control{os.sep}Control.csv')))
loop_time = int(control['Check Frequency in Minutes'])

while True:
    if int(dt.datetime.now().strftime("%H")) == int(
            dt.datetime.strptime(control['Time to Stop'], '%I:%S %p').strftime("%H")):
        break
    try:

        data = getData()

        try:
            checks(data)
        except KeyError:
            pass

        retry_time = dt.datetime.now() + dt.timedelta(minutes=loop_time)
        print(f'Retrying in {retry_time.strftime("%H:%M")} ...')
        pause.until(retry_time)

    except KeyboardInterrupt:

        print('Exiting')
        break
