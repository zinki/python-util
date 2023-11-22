import xlrd


# 将excel表格内容导入到tables列表中
def import_excel(excel):
    header = excel.row_values(0)
    array = {}
    tables = []
    for rown in range(1, excel.nrows):
        for index in range(len(header)):
            array[header[index]] = excel.cell_value(rown, index)
        tables.append(array)
    return tables


def update(arrays, tableName):
    sql = "UPDATE %s SET "
    card_bin = ''
    for key in arrays:
        if key == 'card_bin':
            card_bin = str(arrays[key])
            continue
        value = ''
        if arrays[key] == 'NULL':
            value = 'NULL'
        else:
            value += '\''
            value += str(arrays[key])
            value += '\''
        sql += key
        sql += ' = '
        sql += value
        sql += ','

    sql = sql[:len(sql) - 1]
    sql += ' WHERE card_bin = \''
    sql += card_bin
    sql += '\';'
    sql = sql % tableName
    print(sql)
    return sql


def insert(arrays, tableName):
    sql = "INSERT %s("
    values = ''
    for key in arrays:
        value = ''
        if arrays[key] == 'NULL':
            value = 'NULL'
        else:
            value += '\''
            value += str(arrays[key])
            value += '\''
        sql += key
        sql += ','
        values += value
        values += ','

    sql = sql[:len(sql) - 1]
    sql += ') VALUES('
    values = values[:len(values) - 1]
    sql += values
    sql += ');'
    sql = sql % tableName
    print(sql)
    return sql


def readFile(path):
    # 导入需要读取的第一个Excel表格的路径
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    # 创建一个空列表，存储Excel的数据
    tables = import_excel(table)
    return tables


def writeFile(tableName, tables, operate):
    # 创建一个空列表，存储Excel的数据
    with open('output.txt', 'w') as f:
        for i in tables:
            if operate == 'INSERT':
                sql = insert(i, tableName)
            else:
                sql = update(i, tableName)
            f.write(sql)
            f.write('\n')


if __name__ == '__main__':
    # 将excel表格的内容导入到列表中
    file = readFile(r"D:\Users\qzhang59\Documents\卡bin修改\11.22\新增.xls")
    writeFile("card_bin_detail_infos", file, "UPDATE")
