import re

def aspace_dates(aspace_date):
    date_ymd = re.findall('^creation:\s(\d\d\d\d-\d\d-\d\d)$', aspace_date)
    date_ym = re.findall('^creation:\s(\d\d\d\d-\d\d)$', aspace_date)
    date_y = re.findall('^creation:\s(\d\d\d\d)$', aspace_date)
    date_range = re.findall('creation:\s(.+?--.+)', aspace_date)
    date_multiple = re.findall('^creation:\s.+?;\screation:\s.+?$', aspace_date)

    def month_convert(month):
        if month == '01':
            return 'January'
        elif month == '02':
            return 'February'
        elif month == '03':
            return 'March'
        elif month == '04':
            return 'April'
        elif month == '05':
            return 'May'
        elif month == '06':
            return 'June'
        elif month == '07':
            return 'July'
        elif month == '08':
            return 'August'
        elif month == '09':
            return 'September'
        elif month == '10':
            return 'October'
        elif month == '11':
            return 'November'
        elif month == '12':
            return 'December'

    def ymd_convert(convert_ymd):
        date_year = convert_ymd[0].split('-')[0].strip()
        iso_month = convert_ymd[0].split('-')[1].strip()
        date_month = month_convert(iso_month)
        date_day = int(convert_ymd[0].split('-')[2].strip())
        ymd_converted = '{date_month} {date_day}, {date_year}'.format(date_month = date_month, date_day = date_day, date_year = date_year)
        return ymd_converted

    def ym_convert(convert_ym):
        date_year = convert_ym[0].split('-')[0].strip()
        iso_month = convert_ym[0].split('-')[1].strip()
        date_month = month_convert(iso_month)
        ym_converted = '{date_month} {date_year}'.format(date_month = date_month, date_year = date_year)
        return ym_converted

    def y_convert(convert_y):
        date_year = convert_y[0].strip()
        y_converted = '{date_year}'.format(date_year = date_year)
        return y_converted

    if aspace_date == None:
        date_formatted = ' (undated)'
        return date_formatted

    if len(date_ymd) == 1:
        date_formatted = ymd_convert(date_ymd)
        date_formatted = ', ' + date_formatted
        return date_formatted

    if len(date_ym) == 1:
        date_formatted = ym_convert(date_ym)
        date_formatted = ', ' + date_formatted
        return date_formatted

    if len(date_y) == 1:
        date_formatted = y_convert(date_y)
        date_formatted = ', ' + date_formatted
        return date_formatted

    if len(date_range) == 1:
        date_first = date_range[0].split('--')[0].strip()
        date_first_ymd = re.findall('(\d\d\d\d-\d\d-\d\d)$', date_first)
        date_first_ym = re.findall('(\d\d\d\d-\d\d)$', date_first)
        date_first_y = re.findall('(\d\d\d\d)$', date_first)
        date_second = date_range[0].split('--')[1].strip()
        date_second_ymd = re.findall('(\d\d\d\d-\d\d-\d\d)$', date_second)
        date_second_ym = re.findall('(\d\d\d\d-\d\d)$', date_second)
        date_second_y = re.findall('(\d\d\d\d)$', date_second)

        if len(date_first_ymd) == 1:
            date_first_formatted = ymd_convert(date_first_ymd)

        if len(date_first_ym) == 1:
            date_first_formatted = ym_convert(date_first_ym)

        if len(date_first_y) == 1:
            date_first_formatted = y_convert(date_first_y)

        if len(date_second_ymd) == 1:
            date_second_formatted = ymd_convert(date_second_ymd)

        if len(date_second_ym) == 1:
            date_second_formatted = ym_convert(date_second_ym)

        if len(date_second_y) == 1:
            date_second_formatted = y_convert(date_second_y)

        date_formatted = ', {date_first_formatted} - {date_second_formatted}'.format(date_first_formatted=date_first_formatted, date_second_formatted=date_second_formatted)
        return date_formatted

    if len(date_multiple) == 1:
        date_first = date_multiple[0].split(';')[0].split()[1].strip()
        date_first_ymd = re.findall('(\d\d\d\d-\d\d-\d\d)$', date_first)
        date_first_ym = re.findall('(\d\d\d\d-\d\d)$', date_first)
        date_first_y = re.findall('(\d\d\d\d)$', date_first)
        date_second = date_multiple[0].split(';')[1].split()[1].strip()
        date_second_ymd = re.findall('(\d\d\d\d-\d\d-\d\d)$', date_second)
        date_second_ym = re.findall('(\d\d\d\d-\d\d)$', date_second)
        date_second_y = re.findall('(\d\d\d\d)$', date_second)

        if len(date_first_ymd) == 1:
            date_first_formatted = ymd_convert(date_first_ymd)

        if len(date_first_ym) == 1:
            date_first_formatted = ym_convert(date_first_ym)

        if len(date_first_y) == 1:
            date_first_formatted = y_convert(date_first_y)

        if len(date_second_ymd) == 1:
            date_second_formatted = ymd_convert(date_second_ymd)

        if len(date_second_ym) == 1:
            date_second_formatted = ym_convert(date_second_ym)

        if len(date_second_y) == 1:
            date_second_formatted = y_convert(date_second_y)

        date_formatted = ', {date_first_formatted} and {date_second_formatted}'.format(date_first_formatted=date_first_formatted, date_second_formatted=date_second_formatted)
        return date_formatted
