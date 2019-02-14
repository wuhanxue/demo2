# 该方法为调用xlutils的方法，速度较快，耗时约15秒钟，缺点是不保留样式
from xlrd import open_workbook
from xlutils.copy import copy
import os


# 查找关键字方法
def find_keyword_in_sheet(sheet, keyword):
    for i, cell2 in enumerate(sheet.col_values(2)):
        if cell2:
            if str(cell2).split('.')[0] == str(keyword):
                return i


# 计算超额费用
def get_bonus(total, limit):
    return 0 if total - limit <= 0 else round(total - limit, 2)


# 获取额度和超额费用
def get_limit_and_bonus(raw, base, total):
    limit = 0
    bonus = 0
    if raw == 36 and total < 60:
        limit = 36
        bonus = 0
    elif raw == 36 and base >= 60:
        limit = 63
        bonus = get_bonus(total, limit)
    elif raw == 60 and base == 60:
        limit = 60
        bonus = get_bonus(total, limit)
    elif raw == 60 and 60 < base <= 70:
        limit = 87
        bonus = get_bonus(total, limit)
    elif raw == 60 and base > 70:
        limit = 90
        bonus = get_bonus(total, limit)
    elif raw > 60:
        limit = base
        bonus = get_bonus(total, raw)
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
    file_path_C = file_path_A.split('.')[0] + '_程序生成.xls'
    print('读取文件(%s, %s)...(时间较长，请勿关闭程序)' % (file_path_A, file_path_B))
    # 打开文件
    workbookA = open_workbook(file_path_A)
    workbookB = open_workbook(file_path_B)
    workbookC = copy(workbookA)
    # 猜测格式类型
    workbookA.guess_types = True
    workbookB.guess_types = True
    # 读取表格
    sheetA = workbookA.sheet_by_name('Sheet1')
    sheetB = workbookB.sheet_by_name('Sheet1')
    sheetC = workbookC.get_sheet('Sheet1')
    print('读取完成!')
    print('正在计算...')
    # 循环A表
    A_row = 0
    for cellC, cellD in zip(sheetA.col_values(2), sheetA.col_values(3)):
        # 判断是否为空
        if cellC is None or cellC == '网络':
            break
        if cellC and A_row != 0:
            # 查找电话号码
            number = str(cellC).split('.')[0]  # 电话号码
            find_B_row = find_keyword_in_sheet(sheetB, number)
            sheetC.write(A_row, 2, number)  # 电话
            if find_B_row:
                raw = float(cellD)  # 套餐费
                base = float(sheetB.cell(find_B_row, 3).value)  # 基本
                total = float(sheetB.cell(find_B_row, 14).value)  # 总费用
                limit, bonus = get_limit_and_bonus(raw, base, total)
                # sheetA['E'][A_row].value = raw                        # 额度
                sheetC.write(A_row, 5, raw + bonus)  # 月总费用
                sheetC.write(A_row, 6, bonus)  # 超额费用
                # print(A_row, end='-->')
                # print(limit, bonus)
            else:
                print('%s 未找到!' % str(cellC))
                sheetC.write(A_row, 5, '未找到')  # 月总费用
                sheetC.write(A_row, 6, '未找到')  # 超额费用
        A_row += 1  # 行号加1
    print('计算完成！')
    # 判断文件是否存在，若存在则删除
    if os.path.exists(file_path_C):
        os.remove(file_path_C)
    print('正在保存...')
    workbookC.save(file_path_C)
    print('保存成功！请查看文件 %s' % file_path_C)


# 主函数
if __name__ == '__main__':
    try:
        read_excel()
    except PermissionError as e:
        print('Excel文件被占用了，要先关掉才行')
    except Exception as e:
        print('没有找到文件,需要将扣款和话单文件放在一起')
        print('Error:', e)
    os.system('pause')
