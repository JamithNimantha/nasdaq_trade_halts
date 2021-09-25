import feedparser

dic = {'Halt Date': ['09/20/2021', '02/22/2019'], 'Halt Time': ['19:14:06', '16:02:01'], 'Issue Symbol': ['TNT', 'SVA'],
       'Issue Name': ['Peak Fintech Group Inc. Cm', 'Sinovac Biotech, Ltd Ord Shrs'], 'Market': ['NASDAQ', 'NASDAQ'],
       'Reason Codes': ['T12', 'T12'], 'Pause Threshold Price': ['', ''], 'Resumption Date': ['', ''],
       'Resumption Quote Time': ['', ''], 'Resumption Trade Time': ['', '']}


def getData():
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


getData()
