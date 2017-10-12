import xlrd
path = 'sai_data.xlsx'
workbook = xlrd.open_workbook(path)
worksheet = workbook.sheet_by_index(0)

offset = 0


def X(data):
    try:
        if isinstance(data, (int, float)):
            data = str(data)
        else:
            data = str(data).replace(',', '')
        return ''.join([chr(ord(x)) for x in data])
    except ValueError:
        return data.encode('utf8')

import csv
with open('sai.csv', 'wb') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=',',
                            quotechar=' ')
    titles = ["Name","Age","Adress"]
    spamwriter.writerow(titles)
    rows = []
    for i, row in enumerate(range(worksheet.nrows)):
        csv_row = []
        if i <= offset: 
            continue
        r = []
        for j, col in enumerate(range(worksheet.ncols)):
            r.append(worksheet.cell_value(i, j))
        rows.append(r)
        csv_row.append(r[0])
        csv_row.append(r[1])
        csv_row.append(r[2])
        
        
        print r, type(r)
        r = [X(i) for i in csv_row] 
        spamwriter.writerow(r)  
    print 'Got %d rows' % len(rows)


