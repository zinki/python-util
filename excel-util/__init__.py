import xlrd

# 导入需要读取的第一个Excel表格的路径

data = xlrd.open_workbook(r'D:\Users\qzhang59\Documents\卡bin修改\11.15\卡BIN更新记录20231114.xls')

table = data.sheets()[0]

# 创建一个空列表，存储Excel的数据

tables = []


# 将excel表格内容导入到tables列表中

def import_excel(excel):
    for rown in range(excel.nrows):
        array = {'bank_name': table.cell_value(rown, 0), 'bank_abbr': table.cell_value(rown, 1),
                 'card_bin': int(table.cell_value(rown, 2)), 'card_type': int(table.cell_value(rown, 3)),
                 'card_orgnz': table.cell_value(rown, 4), 'card_name': table.cell_value(rown, 5),
                 'cardno_len': int(table.cell_value(rown, 6)), 'status': int(table.cell_value(rown, 7)),
                 'cardbin_len': int(table.cell_value(rown, 8)), 'country_name': table.cell_value(rown, 9),
                 'primary_scheme': table.cell_value(rown, 10), 'country_2_code': table.cell_value(rown, 11),
                 'country_3_code': table.cell_value(rown, 12), 'country_3_num': int(table.cell_value(rown, 13)),
                 'card_grade': table.cell_value(rown, 14)}

        tables.append(array)


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


def writeFile(tableName, operate):
    import_excel(table)
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
    writeFile("middle_card_bin_infos", "INSERT")