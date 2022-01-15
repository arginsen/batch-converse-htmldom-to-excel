from bs4 import BeautifulSoup
from listener import MsgBoxListener
import xlwings as xw
import json
import math
import time
import re
import os

# 获取当前目录下的 html
html_path = './html-doc/'
xlsb_path = './xlsb-doc/'
output_path = './html-parse-output/'

stat_target = 'statistics-doc.xlsx'

# 收集
html_list = []
xlsb_list = []
stat_list = {}

# 收集表格数据
table_head_list = []
table_body_list = []

# 针对提取出的数据写入目标 excel 特定位置
def editMatchedXlsb(data, filename, doctime, powerclass, snnumber):
    '''
    https://buildmedia.readthedocs.org/media/pdf/xlwings/stable/xlwings.pdf
    '''
    app = xw.App(visible=False,add_book=False)
    #不显示Excel消息框
    app.display_alerts = False
    #关闭屏幕更新,可加快宏的执行速度
    app.screen_updating = False

    wb = app.books.open(xlsb_path + filename + '.xlsb')
    # 也可以直接用 Book 创建，就不用定义 app，之后直接关闭 wb 即可
    # wb = xw.Book('test.xlsb')

    # 处理 [封面] 
    sheet1 = wb.sheets['封面']
    # 楼层
    sheet1.range('Q33').value = 'F1' if filename[0] == '1' else 'F2' if filename[0] == '2' else 'F3' if filename[0] == '3' else 'F4'
    # Site ID
    sheet1.range('Q35').value = '101' if filename[0] == '1' else '201' if filename[0] == '2' else '301' if filename[0] == '3' else '401'
    # 设备编码
    sheet1.range('Q39').value = filename
    # sn 码
    sheet1.range('Q40').value = snnumber
    # 时间
    sheet1.range('E41').value = doctime
    # 保留 stat 数据
    stat_item = []
    stat_item.append(filename) # 0 机器编码
    snnumber_item = []
    if filename[-1] == '1':
        snnumber_item.append(snnumber) # 条码为数组
        stat_item.append(snnumber_item) # 1 产品条码
        stat_item.append('1') # 2 本机编码
        stat_item.append('1') # 3 并机台数
        stat_list[filename] = stat_item
    elif filename[-1] == '2':
        # 结尾为 2 说明并机台数为 2，存在两个条码
        pre_key = filename[:-1] + '1'
        for k, v in stat_list.items():
            if k == pre_key:
                # 给前一个 1 添加，并更改
                stat_list[pre_key][1].append(snnumber)
                stat_list[pre_key][3] = '2'
                # 当前项
                snnumber_item.append(stat_list[pre_key][1][0])
                snnumber_item.append(snnumber)
                stat_item.append(snnumber_item)
                stat_item.append('2')
                stat_item.append('2')
                stat_list[filename] = stat_item
                break
    elif filename[-1] == '3':
        for i in range(2):
            crr_pre_key = filename[:-1] + str(i + 1)
            for k, v in stat_list.items():
                if k == crr_pre_key:
                    # 给前一个 1 添加，并更改
                    stat_list[crr_pre_key][1].append(snnumber)
                    stat_list[crr_pre_key][3] = '3'
                    # 当前项仅编辑一次
                    if i == 0:
                        snnumber_item.append(stat_list[crr_pre_key][1][0])
                        snnumber_item.append(stat_list[crr_pre_key][1][1])
                        snnumber_item.append(snnumber)
                        stat_item.append(snnumber_item)
                        stat_item.append('3')
                        stat_item.append('3')
                        stat_list[filename] = stat_item
                    break

    # 处理 [巡检报告]
    sheet2 = wb.sheets('巡检报告')
    # 功率等级
    sheet2.range('G33').value = powerclass
    # 电池型号
    sheet2.range('G40').value = '南都 6-GFM-180HR' if powerclass == 500 else '南都 6-GFM-155HR'
    # 电池组数
    sheet2.range('G41').value = 2 if powerclass == 300 else 3 if powerclass == 500 else 4
    # 系统类型
    sheet2.range('G35').value = '单机' if powerclass == 300 else '分散并机'
    # 外部维修旁路
    sheet2.range('G36').value = '有-3P'
    # 内部开关配置
    sheet2.range('R33').value = '主路输入、旁路、输出' if powerclass == 600 else '主路输入、旁路、输出、维修旁路'
    # 并机输出开关
    sheet2.range('R44').value = '无'
    # 时间
    sheet2.range('E462').value = doctime
    sheet2.range('Q462').value = doctime
    # 并机/本机编码/产品条码
    # 单独处理

    # Mains
    Mains = data['Mains']
    # 定义一个二维数组存储数据
    Mains_handle = [[0 for y in range(3)] for x in range(3)]
    Mains_index = 0
    # 对数据进行处理
    for k, l in Mains.items():
        for i, v in enumerate(l):
            if v[-2:] == 'Hz':
                Mains_handle[Mains_index][i] = v[:-3].strip()
            elif v[-1:] == 'V':
                Mains_handle[Mains_index][i] = round(float(v[:-1]) * math.sqrt(3), 2)
            else:
                Mains_handle[Mains_index][i] = v[:-1].strip()
        Mains_index += 1
    # 将数据渲染至目标位置 - 纵向
    sheet2.range('G167').options(transpose=True).value = Mains_handle[0]
    sheet2.range('K167').options(transpose=True).value = Mains_handle[1]
    sheet2.range('O167').options(transpose=True).value = Mains_handle[2]
    # sheet2.range('G167:R169').api.HorizontalAlignment = -4131

    # Reserve
    Reserve = data['Reserve']
    Reserve_handle = [[0 for y in range(3)] for x in range(3)]
    Reserve_index = 0
    for k, l in Reserve.items():
        # 清空二维数组初始化元素
        Reserve_handle[Reserve_index].clear()
        for i, v in enumerate(l):
            if v[-1] == 'V':
                Reserve_handle[Reserve_index].append(v[:-1].strip())
            elif v[-2:] == 'Hz':
                Reserve_handle[Reserve_index].append(v[:-2].strip())
        Reserve_index += 1
    sheet2.range('G178').options(transpose=True).value = Reserve_handle[0]
    sheet2.range('K178').options(transpose=True).value = Reserve_handle[1]
    sheet2.range('O178').options(transpose=True).value = Reserve_handle[2]
    # sheet2.range('G178:R179').api.HorizontalAlignment = -4131

    # output
    Output = data['Output']
    Output_handle = [[0 for y in range(9)] for x in range(3)]
    Output_index = 0
    for k, l in Output.items():
        for i, v in enumerate(l):
            # 两个行被清空了 ...
            Output_handle[Output_index][2] = ''
            Output_handle[Output_index][5] = ''
            if v[-1] == 'V':
                Output_handle[Output_index][0] = round(float(v[:-1].strip()) * math.sqrt(3), 2)
                Output_handle[Output_index][1] = v[:-1].strip()
            elif v[-2:] == 'Hz':
                Output_handle[Output_index][3] = v[:-2].strip()
            elif v[-1] == 'A':
                if v[-3:] == 'kVA':
                    Output_handle[Output_index][6] = v[:-3].strip()
                else:
                    Output_handle[Output_index][4] = v[:-1].strip()
            elif v[-2:] == 'PF':
                Output_handle[Output_index][7] = abs(float('0 PF'[:-2].strip()))
            else:
                Output_handle[Output_index][8] = float(l[3][:-3].strip()) * 3 / powerclass
        Output_index += 1
    sheet2.range('G263').options(transpose=True).value = Output_handle[0]
    sheet2.range('K263').options(transpose=True).value = Output_handle[1]
    sheet2.range('O263').options(transpose=True).value = Output_handle[2]
    # sheet2.range('G263:R269').api.HorizontalAlignment = -4131

    # 保存
    wb.save()
    wb.close()
    app.quit()
    time.sleep(0.2)

# 创建 excel 表格并写入 html 解析的数据
def createXlsx(data, filename):
    app = xw.App(visible=False,add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.add()
    sheet = wb.sheets['sheet1']
    # 写入表格数据
    count = 1
    for k, v in data.items():
        '''
        单元格区域合并
        sheet2.range('A1:B1').merge()
        单元格区域清空
        sheet2.range('A1:D3').clear()
        设置字体为粗体
        sheet2.range('A1').api.Font.Bold = True
        水平居中
        sheet2.range('A1').api.HorizontalAlignment = -4108
        单元格格式：http://www.dszhp.com/xlwings-format.html
        '''
        sheet.range('A' + str(count) + ':G' + str(count)).merge()
        sheet.range('A' + str(count)).value = k
        sheet.range('A' + str(count)).api.Font.Bold = True
        sheet.range('A' + str(count)).api.Font.Size = 16 
        count += 1
        if not isinstance(v, list):
            for i in v:
                sheet.range('A' + str(count)).value = i
                sheet.range('A' + str(count)).api.Font.Bold = True
                sheet.range('B' + str(count)).value = v[i]
                count += 1
        else:
            sheet.range('B' + str(count)).value = v
    # 整体居中
    sheet.range('A1:G' + str(count)).api.HorizontalAlignment = -4108

    if not os.path.isdir(output_path[2:-1]):
        os.mkdir(r'' + output_path[2:-1])
    wb.save(output_path + filename + '.xlsx')
    wb.close()
    app.quit()
    time.sleep(0.2)

# 遍历每个 html
def parseHtml(file, filename):
    # 得到整个文档 dom
    doc = BeautifulSoup(open(file), from_encoding='utf-8', features='html.parser')

    # 得到时间
    time_dom = doc.find('h3')
    time_compile = re.compile('\d{4}-\d{2}-\d{2}')
    doctime = re.search(time_compile, str(time_dom))[0].replace('-', '/', 2)

    # 得到功率等级
    power_class_dom = doc.find('table')
    power_class_compile = re.compile(r'<td>(\d+)kVA')
    power_class = re.search(power_class_compile, str(power_class_dom))[1]

    # 得到设备编码 sn
    sn_number_dom = doc.find('h1')
    sn_number_compile = re.compile('\((.+)\)')
    sn_number = re.search(sn_number_compile, str(sn_number_dom))[1]

    # 得到对应 meatures 下的表格节点
    measures = doc.find('h3', string='Measures')
    table = measures.find_parent().find_next_sibling()

    # 获得表头字段
    table_head = table.find('thead').find_all('td')
    for td in table_head:
        if td.string != None:
            table_head_list.append(td.string)

    # 获得表身数据 并整理为一个对象
    table_body = table.find('thead').find_next_sibling().contents
    table_body_list = {}
    table_body_item_flag = 0
    for x in table_body:
        if re.search(re.compile('table'), str(x)):
            item = x.find_all('td')
            item_transform = {}
            # 当前 L1, L2, L3 的计数
            curr_L = []
            L_index = 1
            L_length = 3
            # 当前遍历 td 的长度
            item_length = len(item)
            # 当前遍历 td 内元素的计数
            el_length = item_length // L_length
            el_last_length = item_length - el_length * 2
            el_index = 1
            for v in item:
                v = v.string
                curr_L.append(v)
                # # 对 battery 特殊处理
                if table_body_item_flag == 3:
                    item_transform = curr_L
                elif (el_index == el_length and L_index != 3) or (L_index == 3 and el_index == el_last_length):
                    item_transform['L' + str(L_index)] = curr_L
                    curr_L = []
                    el_index = 0
                    L_index += 1
                el_index += 1
            # 数组拍平 -----
            # item_transform = curr_L
            # -------------
            table_body_list[table_head_list[table_body_item_flag]] = item_transform
            table_body_item_flag += 1
            if table_body_item_flag >= 4:
                break

    # 对表格进行补全
    export_excel = table_body_list.copy()
    for k, v in export_excel.items():
        if not isinstance(v, list):
            v['L1'].append('50 Hz')
            v['L2'].append('50 Hz')
    # 记录当前提取内容
    with open('result.txt', 'wb') as file:
        file.write(json.dumps(table_body_list, indent=4, ensure_ascii=False, sort_keys=True).encode())
    # 解析后的内容创建 excel 记录
    # createXlsx(export_excel, filename)
    # 若解析的文档存在与之对应的待编辑 xlsb 文件
    if filename in xlsb_list:
        editMatchedXlsb(table_body_list, filename, doctime, int(power_class), sn_number)
    else:
        print('提供的 excel 文档中不存在 [' + filename + ']')

# 获取统计数据 stat
def handleStat():
    # 开始新的遍历获得 stat_list
    finished_doc = 1
    for v in xlsb_list:
        app = xw.App(visible=False,add_book=False)
        #不显示Excel消息框
        app.display_alerts = False
        #关闭屏幕更新,可加快宏的执行速度
        app.screen_updating = False

        wb = app.books.open(xlsb_path + v + '.xlsb')
        sheet = wb.sheets['巡检报告']

        # 并机数/本机编号
        sheet.range('R43').value = stat_list[v][2]
        sheet.range('G43').value = stat_list[v][3]
        # 并机条码
        snnumber_range = ['G55', 'R55', 'G56']
        for i, s in enumerate(stat_list[v][1]):
            sheet.range(snnumber_range[i]).value = s
        # 并机检测结果
        sheet.range('T379').value = '正常'
        sheet.range('T381').value = '正常'
        sheet.range('T382').value = '未测试'

        # 导出 pdf
        # wb.to_pdf(r'' + xlsb_path + v + '.pdf', ['封面', '巡检报告'])
        # 退出
        wb.save()
        wb.close()
        app.quit()
        time.sleep(0.2)
        print('[' + v[0] + '] 数据已整理，当前已有' + str(finished_doc) + '个完成')
        finished_doc += 1

def main():
    # 判断目标文件夹是否存在
    if not os.path.isdir(xlsb_path[2:-1]):
        os.mkdir(r'' + xlsb_path[2:-1])
        print('请将代编辑的 excel 文档移入文件夹: ' + xlsb_path[2:-1])
        return
    if not os.path.isdir(html_path[2:-1]):
        os.mkdir(r'' + html_path[2:-1])
        print('请将代编辑的 html 文档移入文件夹: ' + html_path[2:-1])
        return

    # 获取代编辑文件，收集名字为数组，写入后从数组删除
    xlsb_files_list = os.listdir(xlsb_path)
    for i, v in enumerate(xlsb_files_list):
        a, b = os.path.splitext(v)
        if b == ('.xlsb' or '.xlsx'):
            xlsb_list.append(a)

    # 排序处理，之后在获取并机数据时需要按序处理
    xlsb_list.sort()

    # 开启监听
    listener = MsgBoxListener(2)
    listener.start()

    # 抓取 html 数据并写入目标文件
    html_list = os.listdir(html_path)
    finished_doc = 1
    for i in html_list:
        a, b = os.path.splitext(i)
        if b == '.html':
            parseHtml(html_path + i, a)
            print('[' + a + '] 文件已解析，当前已有' + str(finished_doc) + '个完成')
            finished_doc += 1

    # 处理统计文件 statistics-doc
    print("\n")
    print('－－－－－－－－－－－－－－－－－－－－－－－－')
    print('进行数据整合，请等待...')
    handleStat()
    print('处理完毕！')
    listener.stop()

if __name__ == '__main__':
    main()