import xlrd, xlwt
from xlutils.copy import copy
import re
import tkinter as tk
from tkinter import filedialog


# TEST GIT

# 读取"确认表"，创建列表
# ·字典内分别设置 编号、联系电话、家庭住址、家庭成员
# 数据核对：检测sheet名、表格户主名与成员信息中的户主名是否一致，若不一致则报错
#           检测身份证号是否缺失
def read_data(filename):
    info = []  # 储存确认登记表中的信息
    errors = []  # 储存疑似有误的信息
    id_confirm = 0  # 防止编号输入有误

    # 打开文件
    workbook = xlrd.open_workbook(filename)
    sheetnames = workbook.sheet_names()
    for sheetname in sheetnames:

        sheet = workbook.sheet_by_name(sheetname)
        #print(sheetname)
        id_strs = sheet.cell(1, 10).value
        # print(id_strs)
        # print(type(id_strs))
        # print(type(id_strs) == float)
        if (type(id_strs) == float):
            id_strs = str(int(id_strs))
        # if there exist a '/' in id, replace '/' with '-'
        id_strs = id_strs.replace('/', '-', 1);
        id = id_strs.split("-", 1)[0]
        master = sheet.cell(2, 2).value
        Phone_number = sheet.cell(2, 7).value
        address = sheet.cell(3, 2).value

        if len(id_strs.split("-", 1)) == 1 or id_strs.split("-", 1)[1] == '1':
            id_confirm += 1

        '''
        print("id_full:\t", sheet.cell(1, 10).value)
        print("id:\t\t\t", id)
        print("id_confirm:\t", id_confirm)
        '''

        # 根据家庭成员行数获取家庭人数
        headcount = 0
        while sheet.cell(9 + headcount, 0).value and sheet.cell(9 + headcount, 0).value != "家庭\n代表         \n意见":
            headcount += 1
        # while sheet.cell(9 + headcount, 0).value:
        # print(any(char.isdigit() for char in sheet.cell(9+headcount, 0).value))
        # 信息核对
        if re.findall("\d+", sheet.cell(7, 10).value):
            if int(re.findall("\d+", sheet.cell(7, 10).value)[0]) != headcount:
                errors.append("编号为{}的表中家庭人数不一致".format(id))
        else:
            errors.append("编号为{}的表中家庭人数不一致".format(id))

        # 信息核对
        if master.strip() != sheetname.strip():
            errors.append("编号为{}的表中sheet名与户主名不一致。分别为：".format(id))
            errors.append(sheetname.replace(" ", ""))
            errors.append(master.replace(" ", ""))
            # errors.append(master.replace(" ","") != sheetname.replace(" ",""))
            # errors.append(master.strip() != sheetname.strip())
        if (id == ''):
            # print("empty id and id_confirm is:")
            # print(id_confirm)
            errors.append("编号应为{}的表的编号缺少编号".format(id_confirm))
        elif int(id) != id_confirm:
            errors.append("编号为{}的表的编号存在错误".format(id))
            errors.append("id_confirm: {}".format(id_confirm))
        if sheet.cell(9, 2).value != '户主' :
            #print(sheet.cell(9, 2))
            #print(sheet.cell(9, 2) != "户主")

            errors.append("编号为{}的表第10行不为“户主”".format(id))
        # 信息核对

        members = []

        # print("headcount")
        # print(headcount)
        for i in range(int(headcount)):
            row = 9 + int(i)

            id_number = sheet.cell(row, 4).value.strip()
            if id_number == '':
                # 户主无身份证号
                errors.append("编号为{}的表中身份证号缺失".format(id))
                errors.append("缺失行数为：{}".format(row + 1))
                # errors.append(i)
                # errors.append(sheet.cell(row - 1, 0))
                # errors.append(sheet.cell(row, 0))
                gender = ''
            else:
                # print(len(id_number))
                if len(id_number) != 18:
                    errors.append("编号为{}的表中身份证号位数不为18".format(id))
                    errors.append("错误行数为：{}".format(row + 1))
                if id_number[-2] == 'X':
                    print("X occurs")
                    print(i)
                    print(id_number)
                    print(master)
                if id_number[-2].isdigit():
                    if int(id_number[-2]) % 2 == 0:
                        gender = '女'
                    else:
                        gender = '男'
                else:
                    gender = '错'
                    errors.append("编号为{}的表中身份证号倒数第二位不是数字".format(id))
                    errors.append("错误行数为：{}".format(row + 1))
            # 信息核对
            if sheet.cell(row, 2).value.strip() == "户主" and sheet.cell(row, 0).value.strip() != sheetname.strip():
                errors.append("编号为{}的表中sheet名与成员信息的户主名不一致。分别为： ".format(id))
                errors.append(sheetname.strip())
                errors.append(sheet.cell(row, 0).value.strip())
                # errors.append(sheet.cell(row, 0).value.strip()!=sheetname.strip())
            elif sheet.cell(row, 2).value.strip() == "户主" and sheet.cell(row, 4).value.strip() != sheet.cell(4,
                                                                                                             7).value:
                errors.append("编号为{}的表中户主证件号码前后不一致".format(id))

            members.append([sheet.cell(row, 0).value, sheet.cell(row, 2).value, gender, sheet.cell(row, 4).value,
                            sheet.cell(row, 8).value])

        info.append([id, master, Phone_number, address, headcount, members])
        # print(info[0])
    return info, errors


# 根据列表info传入的信息，在原有表格的基础上进行填写
# 从表格的第四行进行填写，若第四行及之后存在内容，则会被覆盖
def write_data(filename, info):
    f = xlwt.Workbook()  # 创建工作簿

    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 进行写操作

    # 设置单元格格式

    # 字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.height = 20 * 11

    # 对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01

    # 边框
    # DASHED虚线、NO_LINE没有、THIN实线
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 样式
    style = xlwt.XFStyle()  # 初始化样式
    style.font = font
    style.alignment = alignment
    style.borders = borders

    # 对齐方式
    alignment2 = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment2.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment2.vert = 0x01

    style2 = xlwt.XFStyle()  # 初始化样式
    style2.font = font
    style2.alignment = alignment2
    style2.borders = borders
    style2.alignment.wrap = 1  # 设置自动换行

    # 当前行数,从第四行开始操作，因此为3
    row_now = 3
    total_menbers = 0

    # 根据info进行信息填写
    for i in range(len(info)):
        # 获取该户总人数，确定需合并的单元格行数
        headcount = int(info[i][4])

        # 填入 序号
        sheet1.write_merge(row_now, row_now + headcount - 1, 0, 0, info[i][0], style)
        # 填入 户主
        sheet1.write_merge(row_now, row_now + headcount - 1, 1, 1, info[i][1], style)
        # 填入 家庭总人口数
        sheet1.write_merge(row_now, row_now + headcount - 1, 2, 2, headcount, style)

        for j in range(headcount):
            # 填入 户内成员姓名
            sheet1.write(row_now + j, 3, info[i][5][j][0], style)
            # 填入 与户主关系
            sheet1.write(row_now + j, 4, info[i][5][j][1], style)
            # 填入 性别
            sheet1.write(row_now + j, 5, info[i][5][j][2], style)
            # 填入 身份证号
            sheet1.write(row_now + j, 6, info[i][5][j][3], style)
            # 填入 证件类型
            sheet1.write(row_now + j, 7, '户口本', style)  # 暂不填入
            # 填入 备注
            sheet1.write(row_now + j, 10, info[i][5][j][4], style)  # 暂不填入
        # 填入 家庭住址
        sheet1.write_merge(row_now, row_now + headcount - 1, 8, 8, info[i][3], style2)
        # 填入 联系电话
        sheet1.write_merge(row_now, row_now + headcount - 1, 9, 9, info[i][2], style)

        total_menbers += headcount
        row_now += headcount  # 根据人数进行移动

    sheet1.write(row_now, 0, "合计", style)
    sheet1.write(row_now, 1, len(info), style)
    sheet1.write(row_now, 2, total_menbers, style)

    for i in range(1000):
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 22  # 设置行高

    for i in range(20):
        sheet1.col(i).width_mismatch = True
    sheet1.col(0).width = 256 * 7
    sheet1.col(1).width = 256 * 11
    sheet1.col(2).width = 256 * 7
    sheet1.col(3).width = 256 * 11
    sheet1.col(4).width = 256 * 11
    sheet1.col(5).width = 256 * 6
    sheet1.col(6).width = 256 * 24
    sheet1.col(7).width = 256 * 12
    sheet1.col(8).width = 256 * 20
    sheet1.col(9).width = 256 * 21
    sheet1.col(10).width = 256 * 17

    f.save(filename)


# 用户界面设计
def GUI():
    root = tk.Tk()

    root.title("文件处理")
    root.geometry('500x550+500+100')

    text1 = tk.Text(root, width=65, height=30)
    text1.insert("end", "只能处理*.xls文件，若打开*.xlsx文件可能会不稳定\n")
    text1.insert("end", "请确保工作簿内没有无用工作表\n")

    scroll = tk.Scrollbar()
    # 将滚动条填充
    scroll.pack(side=tk.RIGHT, fill=tk.Y)  # side是滚动条放置的位置，上下左右。fill是将滚动条沿着y轴填充

    # 将滚动条与文本框关联
    scroll.config(command=text1.yview)  # 将文本框关联到滚动条上，滚动条滑动，文本框跟随滑动
    text1.config(yscrollcommand=scroll.set)  # 将滚动条关联到文本框

    def open_input():
        global fileName_input
        fileName_input = filedialog.askopenfilename(
            filetypes=[("Excel", ".xls"), ("Excel", ".xlsx")])
        text1.insert('end', "\n\n确认表为：" + fileName_input)

        global info, errors
        info, errors = read_data(fileName_input)

        if len(errors) == 0:
            text1.insert("end", "\n\n没有发现错误。\n ")
            return

        text1.insert("end", "\n\n疑似有误的信息如下所示：\n")
        for error in errors:
            text1.insert("end", "\n{}".format(error))


    def open_output():
        global fileName_output
        fileName_output = filedialog.askopenfilename(
            filetypes=[("Excel", ".xls"), ("Excel", ".xlsx"), ("python", ".py")])
        text1.insert("end", "\n\n汇总表为：" + fileName_output)

    def write():
        # strlist = fileName_input.split('.')
        fileName_output = fileName_input[:-4] + '_汇总表.xls'
        write_data(fileName_output, info)
        text1.insert("end", "\n\n填写完成，快去检查一下。\n汇总表为：" + fileName_output)

    def order():
        strlist = fileName_output.split('.')
        filename2 = strlist[0] + '_修改.xls'
        group, info_sorted, errors = reorder(fileName_output)
        write_data_2(group, info_sorted, errors, filename2)
        text1.insert("end", "\n\n填写完成，快去检查一下。\n---------------------------------------\n")
        # print("end", "\n\n原汇总表中以下编号存在问题，请修改")
        # print(errors)

    def readShortGUI():
        global fileName_intput_short
        fileName_intput_short = filedialog.askopenfilename(
            filetypes=[("Excel", ".xls"), ("Excel", ".xlsx"), ("python", ".py")])
        global infoShort, errors_short
        infoShort, errors_short = readShort(fileName_intput_short)
        text1.insert('end', "\n\n汇总表为：" + fileName_intput_short)

    def writeShortGUI():
        strlist = fileName_intput_short.split('.')
        fileName_output = strlist[0] + '_股权.xls'
        writeShort(fileName_output, infoShort)
        text1.insert("end", "\n\n填写完成，快去检查一下。\n折股量化表：" + fileName_output)

    def readReprensentativeGUI():
        global fileName_intput_short
        fileName_intput_short = filedialog.askopenfilename(
            filetypes=[("Excel", ".xls"), ("Excel", ".xlsx"), ("python", ".py")])
        global infoShort, errors_short
        infoShort, errors_short = readRepresentative(fileName_intput_short)
        text1.insert('end', "\n\n汇总表为：" + fileName_intput_short)

    def writeRepresentativeGUI():
        strlist = fileName_intput_short.split('.')
        fileName_output = strlist[0] + '_户代表.xls'
        writeRepresentative(fileName_output, infoShort)
        text1.insert("end", "\n\n填写完成，快去检查一下。\n户代表：" + fileName_output)

    def open_order_short_gui():
        global fileName_input_order_short
        fileName_input_order_short = filedialog.askopenfilename(
            filetypes=[("Excel", ".xls"), ("Excel", ".xlsx"), ("python", ".py")])
        text1.insert("end", "\n\n汇总表为：" + fileName_input_order_short)

    def write_order_short_gui():
        strlist = fileName_input_order_short.split('.')
        fileName_output_order_short = strlist[0] + '_修改.xls'
        group, info_sorted, errors = reorder_short(fileName_input_order_short)
        write_data_short(group, info_sorted, errors, fileName_output_order_short)
        text1.insert("end", "\n\n填写完成，快去检查一下。")

    def one_step_short():
        global fileName_one_step
        fileName_one_step = filedialog.askopenfilename(
            filetypes=[("Excel", ".xls"), ("Excel", ".xlsx"), ("python", ".py")])
        text1.insert("end", "\n\n汇总表为：" + fileName_one_step)
        # todo

    def fang():
        fileName_output = fileName_input[:-4] + '_汇总表.xls'
        write_data_fang(fileName_output, info)
        text1.insert("end", "\n\n填写完成，快去检查一下。\n汇总表为：" + fileName_output)

    tk.Label(root, text="---------------------------汇总表---------------------------").pack()
    tk.Button(root, width=15, height=1, text="打开确认表", command=open_input).pack()
    tk.Button(root, width=15, height=1, text="生成汇总表", command=write).pack()

    '''
    tk.Label(root, text="------------------------汇总表排序---------------------------").pack()
    tk.Button(root, width=15,height=1, text="选择汇总表", command=open_output).pack()
    tk.Button(root, width=15, height=1, text="汇总表排序", command=order).pack()
    '''

    tk.Label(root, text="----------------生成折股量化表------------------").pack()
    tk.Button(root, width=15, height=1, text="选择汇总表", command=readShortGUI).pack()
    tk.Button(root, width=15, height=1, text="获取折股量化表", command=writeShortGUI).pack()

    tk.Label(root, text="----------------生成村民代表名单------------------").pack()
    tk.Button(root, width=15, height=1, text="选择汇总表", command=readReprensentativeGUI).pack()
    tk.Button(root, width=15, height=1, text="获取村民代表名单", command=writeRepresentativeGUI).pack()

    '''
    tk.Label(root, text="--------------农村集体经济组织股权确认登记表 排序------------------").pack()
    tk.Button(root, width=15, height=1, text="选择小表", command=open_order_short_gui).pack()
    tk.Button(root, width=15, height=1, text="小表排序", command=write_order_short_gui).pack()
    

    tk.Label(root, text="-------------------农村集体经济组织成员身份界定确认汇总表(芳）------------------").pack()
    tk.Button(root, width=15, height=1, text="打开确认表", command=open_input).pack()
    tk.Button(root, width=15, height=1, text="生成汇总表", command=write_fang).pack()
    '''

    text1.pack()
    tk.mainloop()


# ------------------------------Order Begin-------------------------------------------------------------

# group 存储id的顺序，info_sorted 存储排序后的信息， errors存储错误
def write_data_2(group, info_sorted, errors, filename2):
    f = xlwt.Workbook()  # 创建工作簿

    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 进行写操作

    # 设置单元格格式

    #  列宽、行高
    # 设置第一列宽度，调整后面的16就好
    sheet1.col(0).width = 256 * 16
    # 设置所有行高度，调整height的值就好。1000是随便写的，如果数据多于1000行就改一下
    tall_style = xlwt.easyxf('font:height 720')
    for i in range(1000):
        # sheet1.row(i).set_style(tall_style)
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24

    # 字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.height = 20 * 11

    # 对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01

    # 边框
    # DASHED虚线、NO_LINE没有、THIN实线
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 样式
    style = xlwt.XFStyle()  # 初始化样式
    style.font = font
    style.alignment = alignment
    style.borders = borders

    # 当前行数,从第四行开始操作，因此为3
    row_now = 3

    id_reorder = 0
    for i in group:
        for inf in info_sorted:

            if i == inf[0]:
                # 开始填写
                id_reorder += 1
                # 获取该户总人数，确定需合并的单元格行数
                headcount = int(inf[2])
                # print(i,headcount)

                # 填入 序号
                sheet1.write_merge(row_now, row_now + headcount - 1, 0, 0, id_reorder, style)
                # 填入 户主
                sheet1.write_merge(row_now, row_now + headcount - 1, 1, 1, inf[1], style)
                # 填入 家庭总人数
                sheet1.write_merge(row_now, row_now + headcount - 1, 2, 2, headcount, style)

                for j in range(headcount):
                    # 填入 户内成员姓名
                    sheet1.write(row_now + j, 3, inf[5][j][0], style)
                    # 填入 与户主关系
                    sheet1.write(row_now + j, 4, inf[5][j][1], style)
                    # 填入 性别
                    sheet1.write(row_now + j, 5, inf[5][j][2], style)
                    # 填入 身份证号
                    sheet1.write(row_now + j, 6, inf[5][j][3], style)
                    # 填入 证件类型
                    sheet1.write(row_now + j, 7, '户口本', style)  # 暂不填入
                    # 填入 备注
                    sheet1.write(row_now + j, 10, inf[5][j][4], style)  # 暂不填入
                # 填入 家庭住址
                sheet1.write_merge(row_now, row_now + headcount - 1, 8, 8, inf[3], style)
                # 填入 联系电话
                sheet1.write_merge(row_now, row_now + headcount - 1, 9, 9, inf[4], style)

                row_now += headcount  # 根据人数进行移动

    f.save(filename2)  # 保存文件


def reorder(filename):
    info = []  # 储存汇总表中的信息

    # 打开文件
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)

    nrows = sheet.nrows  # 获取总行数

    print("sheet.ncols")
    print(sheet.ncols)
    print("sheet.nrows: ", sheet.nrows)
    row_now = 3  # 标记当前行
    while (row_now < nrows):
        if isinstance(sheet.cell(row_now, 2).value, float):
            print("454 row_now: ", row_now)
            print(sheet.cell(row_now, 0).value == '')
            if sheet.cell(row_now, 0).value != '':
                id = int(sheet.cell(row_now, 0).value)
            master = sheet.cell(row_now, 1).value
            headcount = int(sheet.cell(row_now, 2).value)
            address = sheet.cell(row_now, 8).value
            phone = sheet.cell(row_now, 9).value
            # note = sheet.cell(row_now, 10).value

            members = []
            for hc in range(headcount):
                # print("row_now + hc")
                # print(row_now + hc)

                member_name = sheet.cell(row_now + hc, 3).value
                member_relation = sheet.cell(row_now + hc, 4).value
                member_gender = sheet.cell(row_now + hc, 5).value
                member_id_number = sheet.cell(row_now + hc, 6).value

                # print("member_name")
                # print(member_name)

                if sheet.ncols >= 11:
                    member_note = sheet.cell(row_now + hc, 10).value
                else:
                    member_note = ''
                members.append([member_name, member_relation, member_gender, member_id_number, member_note])

            info.append([id, master, headcount, address, phone, members])

        row_now += 1

    info_sorted = sorted(info, key=(lambda x: x[2]), reverse=True)

    group = []  # 用来存储为一组的序号

    for i in range(50):
        if i == 0:
            # 这里调整第一页人数
            target = 18
        else:
            # 这里调整每页人数
            target = 22

        sum = 0
        for e in info_sorted:
            if e[0] not in group:
                if sum + e[2] < target:
                    sum += e[2]
                    group.append(e[0])
                elif sum + e[2] > target:
                    continue
                elif sum + e[2] == target:
                    group.append(e[0])
                    break

    # print(group)
    # print(len(group))

    errors = []
    for i in range(1, 135):
        if i not in group:
            # print(i)
            errors.append(i)

    print("\n\nFunction recoder is done!")

    return group, info_sorted, errors


# ------------------------------Order End-------------------------------------------------------------

# ---------------------------农村集体经济组织股权确认登记表 Begin------------------------------------------

def readShort(filename):
    infoShort = []  # 储存汇总表中的信息
    errors_short = []
    # 打开文件
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)

    nrows = sheet.nrows  # 获取总行数
    # 577 nrow: ", nrows)
    row_now = 3  # 标记当前行
    # ignore the last row of the sheet since the row is for summary
    while (row_now < nrows):


        #print("test 581 row_now", row_now)
        #print("test 581 sheet.cell(row_now, 0).value",sheet.cell(row_now, 0).value)
        if (sheet.cell(row_now, 0).value != '') and (not is_number(sheet.cell(row_now, 0).value)):
            break

        '''
        if sheet.cell(row_now, 0).value == '合计' or sheet.cell(row_now, 0).value == '汇总':
            break
        '''
        if isinstance(sheet.cell(row_now, 2).value, float):
            id = ''
            if sheet.cell(row_now, 0) == '':
                errors_short.append("异常：序号缺失")
            else:
                try:
                    id = int(sheet.cell(row_now, 0).value)
                except:
                    errors_short.append("序号不为整数")

            master = sheet.cell(row_now, 1).value
            headcount = int(sheet.cell(row_now, 2).value)
            address = sheet.cell(row_now, 8).value
            # phone = sheet.cell(row_now, 9).value
            note = sheet.cell(row_now, 10).value

            members = []

            for hc in range(headcount):
                #print("test 601 row_now: ", row_now)
                #print("test 602 hc: ", hc)
                try:
                    member_name = sheet.cell(row_now + hc, 3).value
                    member_relation = sheet.cell(row_now + hc, 4).value
                    # member_gender = sheet.cell(row_now + hc, 5).value
                    # member_id_number = sheet.cell(row_now + hc, 6).value
                    member_note = sheet.cell(row_now + hc, 10).value
                    members.append([member_name, member_relation, note])
                except IndexError:
                    print("error row_now + hc: ", row_now + hc)

                # print("test 604 member_relation: ", sheet.cell(row_now + hc, 4).value)
            infoShort.append([id, master, headcount, members, address])
        row_now += 1

    return infoShort, errors_short


def writeShort(filename, infoShort):
    f = xlwt.Workbook()  # 创建工作簿

    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 进行写操作

    # 设置单元格格式

    #  列宽、行高
    # 设置第一列宽度，调整后面的16就好
    sheet1.col(0).width = 256 * 16
    # 设置所有行高度，调整height的值就好。1000是随便写的，如果数据多于1000行就改一下
    tall_style = xlwt.easyxf('font:height 720')
    for i in range(1000):
        # sheet1.row(i).set_style(tall_style)
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24

    # 字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.height = 20 * 11

    # 对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01

    # 边框
    # DASHED虚线、NO_LINE没有、THIN实线
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 样式
    style = xlwt.XFStyle()  # 初始化样式
    style.font = font
    style.alignment = alignment
    style.borders = borders

    # 对齐方式
    alignment2 = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment2.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment2.vert = 0x01

    style2 = xlwt.XFStyle()  # 初始化样式
    style2.font = font
    style2.alignment = alignment2
    style2.borders = borders
    style2.alignment.wrap = 1  # 设置自动换行

    # 根据infoShort进行信息填写

    row_now = 5
    total_menbers = 0

    for i in range(len(infoShort)):
        # 获取该户总人数，确定需合并的单元格行数

        headcount = int(infoShort[i][2])

        # 序号
        sheet1.write_merge(row_now, row_now + headcount - 1, 0, 0, infoShort[i][0], style)
        # 户主
        sheet1.write_merge(row_now, row_now + headcount - 1, 1, 1, infoShort[i][1], style)
        # 每户股数合计
        sheet1.write_merge(row_now, row_now + headcount - 1, 5, 5, 10 * headcount, style)
        # 地址
        sheet1.write_merge(row_now, row_now + headcount - 1, 7, 7, infoShort[i][4], style2)
        for j in range(headcount):
            # 姓名
            sheet1.write(row_now + j, 2, infoShort[i][3][j][0], style)
            # 与户主关系
            sheet1.write(row_now + j, 3, infoShort[i][3][j][1], style)
            # 股数
            sheet1.write(row_now + j, 4, 10, style)
            # 股权类型
            sheet1.write(row_now + j, 6, '', style)
            # 备注
            sheet1.write(row_now + j, 8, infoShort[i][3][j][2], style)

        total_menbers += headcount
        row_now += headcount  # 根据人数进行移动

    sheet1.write(row_now, 0, "合计", style)
    sheet1.write(row_now, 1, len(infoShort), style)
    sheet1.write(row_now, 2, total_menbers, style)

    for i in range(1000):
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24  # 设置行高
    # 设置列宽
    for i in range(20):
        sheet1.col(i).width_mismatch = True
    # 家庭地址
    sheet1.col(7).width = 256 * 23
    # 备注
    sheet1.col(8).width = 256 * 16
    f.save(filename)

#-----------------------------户代表名单 begin------------------------------------------
def readRepresentative(filename):
    infoShort = []  # 储存汇总表中的信息
    errors_short = []
    # 打开文件
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)

    nrows = sheet.nrows  # 获取总行数
    # 577 nrow: ", nrows)
    row_now = 3  # 标记当前行
    # ignore the last row of the sheet since the row is for summary
    while (row_now < nrows):


        #print("test 581 row_now", row_now)
        #print("test 581 sheet.cell(row_now, 0).value",sheet.cell(row_now, 0).value)
        if (sheet.cell(row_now, 0).value != '') and (not is_number(sheet.cell(row_now, 0).value)):
            break

        '''
        if sheet.cell(row_now, 0).value == '合计' or sheet.cell(row_now, 0).value == '汇总':
            break
        '''
        if isinstance(sheet.cell(row_now, 2).value, float):
            id = ''
            if sheet.cell(row_now, 0) == '':
                errors_short.append("异常：序号缺失")
            else:
                try:
                    id = int(sheet.cell(row_now, 0).value)
                except:
                    errors_short.append("序号不为整数")

            master = sheet.cell(row_now, 1).value
            #headcount = int(sheet.cell(row_now, 2).value)
            #address = sheet.cell(row_now, 8).value
            # phone = sheet.cell(row_now, 9).value
            #note = sheet.cell(row_now, 10).value
            gender = sheet.cell(row_now, 5).value
            #members = []

            '''
            for hc in range(headcount):
                #print("test 601 row_now: ", row_now)
                #print("test 602 hc: ", hc)
                try:
                    member_name = sheet.cell(row_now + hc, 3).value
                    member_relation = sheet.cell(row_now + hc, 4).value
                    # member_gender = sheet.cell(row_now + hc, 5).value
                    # member_id_number = sheet.cell(row_now + hc, 6).value
                    member_note = sheet.cell(row_now + hc, 10).value
                    members.append([member_name, member_relation, note])
                except IndexError:
                    print("error row_now + hc: ", row_now + hc)

                # print("test 604 member_relation: ", sheet.cell(row_now + hc, 4).value)
            '''
            infoShort.append([id, master, gender])
        row_now += 1

    return infoShort, errors_short


def writeRepresentative(filename, infoShort):
    f = xlwt.Workbook()  # 创建工作簿

    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 进行写操作

    # 设置单元格格式

    #  列宽、行高
    # 设置第一列宽度，调整后面的16就好
    sheet1.col(0).width = 256 * 16
    # 设置所有行高度，调整height的值就好。1000是随便写的，如果数据多于1000行就改一下
    tall_style = xlwt.easyxf('font:height 720')
    for i in range(1000):
        # sheet1.row(i).set_style(tall_style)
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24

    # 字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.height = 20 * 11

    # 对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01

    # 边框
    # DASHED虚线、NO_LINE没有、THIN实线
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 样式
    style = xlwt.XFStyle()  # 初始化样式
    style.font = font
    style.alignment = alignment
    style.borders = borders

    # 对齐方式
    alignment2 = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment2.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment2.vert = 0x01

    style2 = xlwt.XFStyle()  # 初始化样式
    style2.font = font
    style2.alignment = alignment2
    style2.borders = borders
    style2.alignment.wrap = 1  # 设置自动换行

    # 根据infoShort进行信息填写

    row_now = 0
    total_menbers = 0

    for i in range(len(infoShort)):
        # 获取该户总人数，确定需合并的单元格行数

        #headcount = int(infoShort[i][2])

        if i < 24:
            # 序号
            sheet1.write(row_now, 0, infoShort[i][0], style)
            # 户主
            sheet1.write(row_now, 1, infoShort[i][1], style)
            # 性别
            sheet1.write(row_now, 2, infoShort[i][2], style)
            #备注 设为空
            sheet1.write(row_now, 3, '', style)
        elif i < 48:
            # 序号
            sheet1.write(row_now - 24, 4, infoShort[i][0], style)
            # 户主
            sheet1.write(row_now - 24, 5, infoShort[i][1], style)
            # 性别
            sheet1.write(row_now - 24, 6, infoShort[i][2], style)
            # 备注 设为空
            sheet1.write(row_now - 24, 7, '', style)
        else:
            x = int((i - 48) / 26)
            # 序号
            sheet1.write(row_now - 48 - 26 * x, 8 + 4 * x, infoShort[i][0], style)
            # 户主
            sheet1.write(row_now - 48 - 26 * x, 9 + 4 * x, infoShort[i][1], style)
            # 性别
            sheet1.write(row_now - 48 - 26 * x, 10 + 4 * x, infoShort[i][2], style)
            # 备注 设为空
            sheet1.write(row_now - 48 - 26 * x, 11 + 4 * x, '', style)
        row_now += 1  # 根据人数进行移动

    '''
    sheet1.write(row_now, 0, "合计", style)
    sheet1.write(row_now, 1, len(infoShort), style)
    sheet1.write(row_now, 2, total_menbers, style)
    '''
    '''
    for i in range(1000):
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24  # 设置行高
    # 设置列宽
    for i in range(20):
        sheet1.col(i).width_mismatch = True
    # 家庭地址
    sheet1.col(7).width = 256 * 23
    # 备注
    sheet1.col(8).width = 256 * 16
    '''
    f.save(filename)
#----------------------------户代表名单 end------------------------------------------
# -----------------------------------------------------------------------------------------------

# ------------------------------Order Short Begin-------------------------------------------------------------

# group 存储id的顺序，info_sorted 存储排序后的信息， errors存储错误
def write_data_short(group, info_sorted, errors, filename2):
    f = xlwt.Workbook()  # 创建工作簿

    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 进行写操作

    # 设置单元格格式

    #  列宽、行高
    # 设置第一列宽度，调整后面的16就好
    sheet1.col(0).width = 256 * 16
    # 设置所有行高度，调整height的值就好。1000是随便写的，如果数据多于1000行就改一下
    tall_style = xlwt.easyxf('font:height 720')
    for i in range(1000):
        # sheet1.row(i).set_style(tall_style)
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24

    # 字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.height = 20 * 11

    # 对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01

    # 边框
    # DASHED虚线、NO_LINE没有、THIN实线
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 样式
    style = xlwt.XFStyle()  # 初始化样式
    style.font = font
    style.alignment = alignment
    style.borders = borders

    # 当前行数,从第6行开始操作，因此为5
    row_now = 5
    id_reorder = 0
    for i in group:
        for inf in info_sorted:

            if i == inf[0]:
                # 开始填写
                id_reorder += 1
                # 获取该户总人数，确定需合并的单元格行数
                headcount = len(inf[2])
                # print("headcount")
                # print(i, headcount)

                # 填入 序号
                sheet1.write_merge(row_now, row_now + headcount - 1, 0, 0, id_reorder, style)
                # 填入 户主
                sheet1.write_merge(row_now, row_now + headcount - 1, 1, 1, inf[1], style)
                # 填入 家庭总人数
                # sheet1.write_merge(row_now, row_now + headcount - 1, 2, 2, headcount, style)

                for j in range(headcount):
                    # 填入 户内成员姓名
                    sheet1.write(row_now + j, 2, inf[2][j][0], style)
                    # 填入 与户主关系
                    sheet1.write(row_now + j, 3, inf[2][j][1], style)
                    # 填入 性别
                    # sheet1.write(row_now + j, 5, inf[5][j][2], style)
                    # 填入 身份证号
                    # sheet1.write(row_now + j, 6, inf[5][j][3], style)
                    # 填入 证件类型
                    # sheet1.write(row_now + j, 7, '户口本', style)  # 暂不填入
                    # 填入 备注
                    # sheet1.write(row_now + j, 10, inf[5][j][4], style)  # 暂不填入
                # 填入 家庭住址
                # sheet1.write_merge(row_now, row_now + headcount - 1, 8, 8, inf[3], style)
                # 填入 联系电话
                # sheet1.write_merge(row_now, row_now + headcount - 1, 9, 9, inf[4], style)

                row_now += headcount  # 根据人数进行移动

    f.save(filename2)  # 保存文件


def reorder_short(filename):
    info = []  # 储存汇总表中的信息

    # 打开文件
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)

    nrows = sheet.nrows  # 获取总行数

    row_now = 5  # 标记当前行
    headcount = 0
    while (row_now < nrows):
        # if isinstance(sheet.cell(row_now, 2).value,float):

        # id = int(sheet.cell(row_now, 0).value)
        # master = sheet.cell(row_now, 1).value
        # headcount = int(sheet.cell(row_now, 2).value)
        # address = sheet.cell(row_now, 8).value
        # phone = sheet.cell(row_now, 9).value
        # note = sheet.cell(row_now, 10).value

        # members=[]
        if sheet.cell(row_now, 0).value != '':
            id = int(sheet.cell(row_now, 0).value)
            master = sheet.cell(row_now, 1).value

            last_id = id
            last_master = master

            if id != 1:
                info.append([last_id, last_master, members, headcount])
                headcount = 0

            members = []

        if sheet.cell(row_now, 2) != '' and sheet.cell(row_now, 3):
            member_name = sheet.cell(row_now, 2).value
            member_relation = sheet.cell(row_now, 3).value
            members.append([member_name, member_relation])
            headcount += 1

        row_now += 1

    info.append([id, master, members, headcount])

    info_sorted = sorted(info, key=(lambda x: x[2]), reverse=True)

    group = []  # 用来存储为一组的序号

    for i in range(50):
        if i == 0:
            target = 18
        else:
            target = 22

        sum = 0
        for e in info_sorted:
            if e[0] not in group:
                if sum + e[3] < target:
                    sum += e[3]
                    group.append(e[0])
                elif sum + e[3] > target:
                    continue
                elif sum + e[3] == target:
                    group.append(e[0])
                    break

    # print(group)
    # print(len(group))

    errors = []
    for i in range(1, 135):
        if i not in group:
            # print(i)
            errors.append(i)

    print("\n\nFunction recoder is done!")

    print("info_sorted")
    print(info_sorted)
    return group, info_sorted, errors


# ------------------------------Order Short End-------------------------------------------------------------

# ------------------------------One Step Short Begin------------------------------------------------------


# ------------------------------One Step Short End------------------------------------------------------


# ---------------------------------农村集体经济组织成员身份界定确认汇总表 芳 Begin-----------------------------------------
def write_data_fang(filename, info):
    f = xlwt.Workbook()  # 创建工作簿

    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

    # 进行写操作

    # 设置单元格格式

    # 字体
    font = xlwt.Font()  # 为样式创建字体
    font.name = '宋体'
    font.height = 20 * 11

    # 对齐方式
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01

    # 边框
    # DASHED虚线、NO_LINE没有、THIN实线
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 样式
    style = xlwt.XFStyle()  # 初始化样式
    style.font = font
    style.alignment = alignment
    style.borders = borders

    # 当前行数,从第四行开始操作，因此为3
    row_now = 3

    # 根据info进行信息填写
    for i in range(len(info)):
        # 获取该户总人数，确定需合并的单元格行数
        headcount = int(info[i][4])

        # 填入 序号
        sheet1.write_merge(row_now, row_now + headcount - 1, 0, 0, info[i][0], style)
        # 填入 户主
        sheet1.write_merge(row_now, row_now + headcount - 1, 1, 1, info[i][1], style)
        # 填入 家庭总人口数
        sheet1.write_merge(row_now, row_now + headcount - 1, 2, 2, headcount, style)

        for j in range(headcount):
            # 填入 户内成员姓名
            sheet1.write(row_now + j, 3, info[i][5][j][0], style)
            # 填入 与户主关系
            sheet1.write(row_now + j, 4, info[i][5][j][1], style)
            # 填入 性别
            sheet1.write(row_now + j, 5, info[i][5][j][2], style)
            # 填入 身份证号
            sheet1.write(row_now + j, 6, info[i][5][j][3], style)
            # 填入 证件类型
            sheet1.write(row_now + j, 7, '户口本', style)  # 暂不填入
            # 填入 备注
            sheet1.write(row_now + j, 10, info[i][5][j][4], style)  # 暂不填入
        # 填入 家庭住址
        sheet1.write_merge(row_now, row_now + headcount - 1, 8, 8, info[i][3], style)
        # 填入 联系电话
        sheet1.write_merge(row_now, row_now + headcount - 1, 9, 9, info[i][2], style)

        row_now += headcount  # 根据人数进行移动

    for i in range(1000):
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 20 * 24  # 设置行高

    f.save(filename)


# ---------------------------------农村集体经济组织成员身份界定确认汇总表 芳 End-----------------------------------------

# to judge where a string is a number
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


if __name__ == '__main__':
    # 注意事项
    # 在运行程序前，必须先创建汇总表，且需要确保汇总表处于关闭状态
    # 尽量处理*.xls文件，若打开*.xlsx文件可能会出错，尤其是写入时
    # 程序会从汇总表的第四行开始填写，第四行后存在的内容将会被覆盖
    # 当某一户家庭的户主和sheet名以及成员信息中的户主名不一致时，将会以“有误信息”形式展现
    # 该程序可能存在漏洞，使用后还需多次检查结果

    # 正式代码

    # 若不想使用GUI，则运行以下代码
    # 当需要debug时，使用以下代码更方便
    #  filename1='C:\Users\PJ64\Documents\WeChat Files\hyk460\FileStorage\File\2020-06\海仔村委会白芒村85户确认表-陈雯华.xls'
    #  filename2='C:\Users\PJ64\Desktop\汇总表.xls'
    #  info,errors = read_data(filename1)
    #  write_data(filename2,info)
    #  for error in errors:
    #     print(error)

    # 使用GUI时的代码
    GUI()

    # headcount = re.findall("\d+", '共   人')[0]
    #
    # print(headcount)

    print("\n\nDone!")
