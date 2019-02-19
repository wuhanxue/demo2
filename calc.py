# 该方法为调用openpyxl的方法，速度较慢，耗时约3分钟，但是会保留样式
from openpyxl import load_workbook
import os


# 查找关键字方法
def find_keyword_in_sheet(sheet, keyword):
    for i, cell2 in enumerate(sheet['C']):
        if cell2:
            if str(cell2.value) == str(keyword):
                return i


# 计算超额费用
def get_bonus(total, limit):
    return 0 if total - limit <= 0 else round(total - limit, 2)


# 获取额度和超额费用
def get_limit_and_bonus(raw, raw_limit, base, total):
    limit = 0
    bonus = 0
    if raw == 36 and total < 60:
        limit = 36
        bonus = 0
    elif raw == 36 and base >= 60:
        limit = 63
        if raw_limit > limit:
            limit = raw_limit
        bonus = get_bonus(total, limit)
    elif raw == 60 and base == 60:
        limit = 60
        if raw_limit > limit:
            limit = raw_limit
        bonus = get_bonus(total, limit)
    elif raw == 60 and 60 < base <= 70:
        limit = 87
        if raw_limit > limit:
            limit = raw_limit
        bonus = get_bonus(total, limit)
    elif raw == 60 and base > 70:
        limit = 90
        if raw_limit > limit:
            limit = raw_limit
        bonus = get_bonus(total, limit)
    elif raw > 60:
        limit = base
        if raw_limit > limit:
            limit = raw_limit
        bonus = get_bonus(total, limit)
    return limit, bonus


# 读取Excel
def read_excel():
    file_path_A = ''
    file_path_B = ''
    # 获取文件路径
    # 利用 os 模块中的 listdir 函数，将路径中的所有文件存储到一个 list 变量中。
    files = os.listdir('.')
    # 利用 for 语句浏览 list 变量中的所有元素
    for f in files:
        # 利用 if 语句判断文件名是否符合要求。其中， endswitch 函数用来判断一个字符串是否包含某个后缀。
        # 成员运算符 in 用来判断一个字符串是否包含某个子串。不同的条件用 and 或者 or 来连接。
        if file_path_A != '' and file_path_B != '':
            break
        if '扣款' in f:
            print('找到 ' + f)
            file_path_A = f
        if '话单' in f:
            print('找到 ' + f)
            file_path_B = f
    file_path_C = file_path_A.split('.')[0] + '_程序生成.xlsx'
    print('读取文件(%s, %s)...(时间较长，请勿关闭程序)' % (file_path_A, file_path_B))
    # 打开文件
    workbookA = load_workbook(file_path_A)
    workbookB = load_workbook(file_path_B)
    # 猜测格式类型
    workbookA.guess_types = True
    workbookB.guess_types = True
    # 读取表格
    sheetA = workbookA.active
    sheetB = workbookB.active
    print('读取完成!')
    print('正在计算...')
    # 循环A表
    A_row = 0
    for cellC, cellD, cellE in zip(sheetA['C'], sheetA['D'], sheetA['E']):
        if cellC and A_row != 0:
            # 判断是否为空
            if cellC.value is None:
                break
            # 查找电话号码
            if cellC.value:
                find_B_row = find_keyword_in_sheet(sheetB, cellC.value)
                if find_B_row:
                    raw = cellD.value  # 套餐费
                    raw_limit = int(cellE.value) if str(cellE.value).isdigit() else 0
                    number = str(sheetB['C'][find_B_row].value)  # 电话号码
                    if number == '18112360021':
                        a = '12'
                    base = sheetB['D'][find_B_row].value  # 基本
                    total = round(sheetB['O'][find_B_row].value, 2)  # 总费用
                    limit, bonus = get_limit_and_bonus(raw, raw_limit, base, total)
                    # sheetA['E'][A_row].value = raw                  # 额度
                    if raw >= 100 and raw_limit != 0:
                        if raw < total < raw_limit:
                            sheetA['F'][A_row].value = total
                        else:
                            sheetA['F'][A_row].value = raw_limit + bonus  # 月总费用
                    else:
                        if raw < total < raw_limit:
                            sheetA['F'][A_row].value = total
                        else:
                            sheetA['F'][A_row].value = raw + bonus  # 月总费用
                    sheetA['G'][A_row].value = bonus  # 超额费用
                    # print(A_row, end='-->')
                    # print(limit, bonus)
                else:
                    print('%s 未找到!' % str(cellC.value))
                    sheetA['F'][A_row].value = '未找到'  # 月总费用
                    sheetA['G'][A_row].value = '未找到'  # 超额费用
        A_row += 1  # 行号加1
    print('计算完成！')
    # 判断文件是否存在，若存在则删除
    if os.path.exists(file_path_C):
        os.remove(file_path_C)
    print('正在保存......(时间较长，请勿关闭程序)')
    workbookA.save(file_path_C)
    print('保存成功！请查看文件 %s' % file_path_C)


# 主函数
if __name__ == '__main__':
    try:
        read_excel()
    except PermissionError as e:
        print('Excel文件被占用了，要先关掉才行')
    except Exception as e:
        print('没有找到文件,需要将‘扣款’和‘话单’文件放在一起')
    os.system('pause')
