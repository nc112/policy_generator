import xlwings as xw
from common import translate_numbers_to_words
from json import load as json_load

#TODO: move to data file
policy_dic = {
    "数据产品名称": "",
    "规则名称": "",
    "规则内容": "",
    "字段名": "",
    "运算符": "",
    "规则变量阈值": 0,
    "func_str": ""
}

# TODO: move to data file
personal_data = {
'id': '',
'als_d15_id_pdl_allnum': '',
'als_d15_id_caon_allnum': '',
'als_d15_id_nbank_allnum':'',
'als_d15_id_nbank_nsloan_allnum':'', 
'als_d15_cell_pdl_allnum':'',
'als_d15_cell_caon_allnum':'',
'als_d15_cell_nbank_allnum':'',
'als_d15_cell_nbank_nsloan_allnum':'', 
'als_m1_id_pdl_allnum':'',
'als_m1_id_caon_allnum':'',
'als_m1_id_nbank_allnum':'',
'als_m1_id_nbank_nsloan_allnum':'', 
'als_m1_cell_pdl_allnum':'',
'als_m1_cell_caon_allnum':'',
'als_m1_cell_nbank_allnum':'',
'als_m1_cell_nbank_nsloan_allnum':'', 
'als_m3_id_pdl_allnum':'',
'als_m3_id_caon_allnum':'',
'als_m3_id_nbank_allnum':'',
'als_m3_id_nbank_nsloan_allnum':'', 
'als_m3_cell_pdl_allnum':'',
'als_m3_cell_caon_allnum':'',
'als_m3_cell_nbank_allnum':'',
'als_m3_cell_nbank_nsloan_allnum':'', 
'als_m6_id_pdl_allnum':'',
'als_m6_id_caon_allnum':'',
'als_m6_id_nbank_allnum':'',
'als_m6_id_nbank_nsloan_allnum':'', 
'als_m6_cell_pdl_allnum':'',
'als_m6_cell_caon_allnum':'',
'als_m6_cell_nbank_allnum':'',
'als_m6_cell_nbank_nsloan_allnum':'', 
'als_m12_id_pdl_allnum':'',
'als_m12_id_caon_allnum':'',
'als_m12_id_nbank_allnum':'',
'als_m12_id_nbank_nsloan_allnum':'', 
'als_m12_cell_pdl_allnum':'',
'als_m12_cell_caon_allnum':'',
'als_m12_cell_nbank_allnum':'',
'als_m12_cell_nbank_nsloan_allnum':''
}


def read_person_data_list(person_data_excel):
    assert type(person_data_excel) is str
    assert person_data_excel != ''

    # load configure
    with open('./config.json', 'r') as config_file:
        config_params = json_load(config_file)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    wk = app.books.open(person_data_excel)
    sheet = xw.sheets.active
    last_cell = sheet.used_range.last_cell
    last_translate_col = translate_numbers_to_words(last_cell.column)

    # generate active value range, e.g. A1:F50
    active_cell_range = config_params['start_pos'] + ":" + last_translate_col + str(last_cell.row)
    # print('最大行数:', last_cell.row, "最大列数:", last_translate_col,\
    #         '当前选定sheet活跃区域（有效区域）', active_cell_range)
    data_row_list = sheet.range(active_cell_range).value
    
    personal_data_list = []
    for row_index in range(1, len(data_row_list)):
        index = 0
        for key in personal_data.keys():
            personal_data[key] = data_row_list[row_index][index]
            index += 1
        personal_data_list.append(personal_data.copy())
    #print(personal_data_list)
    return personal_data_list


def read_excel_data(excel_name, print_necessary=False, to_dic=False):
    assert type(excel_name) is str
    assert excel_name != ''

    # load configure
    with open('./config.json', 'r') as config_file:
        config_params = json_load(config_file)

    # fetch active rows and columns
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    wb = app.books.open(excel_name)
    sheet = xw.sheets.active
    last_cell = sheet.used_range.last_cell
    
    last_translate_col = translate_numbers_to_words(last_cell.column)

    # generate active value range, e.g. A1:F50
    range_active_cell = config_params['start_pos'] + ":" + last_translate_col + str(last_cell.row)
    # print('最大行数:', last_cell.row, "最大列数:", last_translate_col,\
    #         '当前选定sheet活跃区域（有效区域）', range_active_cell)
    data_row_list = sheet.range(range_active_cell).value

    # trim data_row_list, replace None to exactly value
    for pos in range(1, len(data_row_list)):
            # collect data
            data_row = data_row_list[pos]
            for index in range(0, len(data_row)):
                if data_row[index] == None:
                    data_row[index] = pre_data_row[index]
            pre_data_row = data_row.copy()
        
    # generate policy params list
    policy_list = []
    global policy_dic
    for pos in range(1, len(data_row_list)):
        index = 0
        for key in policy_dic.keys():
            if key!='func_str':
                policy_dic[key] = data_row_list[pos][index]
                index += 1
            policy_dic['func_str'] = 'result:' + 'result' + policy_dic['运算符'] + str(policy_dic['规则变量阈值'])
            policy_dic['func_str'] = 'lambda ' + str(policy_dic['func_str']).replace(' ','')
        #print(policy_dic['func_str'])
        policy_list.append(policy_dic.copy())
    
    if print_necessary:
        for item in policy_list:
            for key in item.keys():
                print(key, ':', item[key])
    
    # close excel book
    wb.app.kill()

    # return value
    return policy_list