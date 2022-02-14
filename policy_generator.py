import xlwings as xw
from common import translate_table
from json import dump as json_dump
from json import load as json_load


def generate_policy_list_from_xls(excel_name, is_generate_json=False):
    assert type(excel_name) is str
    assert excel_name != ''

    # load configure
    with open('./config.json', 'r') as config_file:
        config_params = json_load(config_file)
    
    start_pos_of_value = config_params['data_start_row']
    policy_column = config_params['字段名']
    operational_character_column = config_params['运算符']
    threshold_column = config_params['规则变量阈值']
    

    # fetch active rows and columns
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(excel_name)
    sheet = xw.sheets.active

    # read policy value list
    #last_translate_col = translate_numbers_to_words(sheet.used_range.last_cell.column)
    range_of_policy = policy_column + start_pos_of_value + ':' + \
                        policy_column + str(sheet.used_range.last_cell.row)
    policy_dic = {}
    for item in sheet.range(range_of_policy).value:
        policy_dic[item] = ''

    # write policy into json file
    if is_generate_json is True:
        with open('policy.json', 'w') as dump_file:
            json_dump(fp=dump_file, obj=policy_dic)

    # close excel book
    wb.app.kill()

    return policy_dic


def policy_soaking(input_list, policy_list):
    assert input_list != ''
    assert policy_list != ''

    #TODO: update here, for dic match algorithm
    with open(r'C:\works\codes\Python\play_ground\file_out.txt', 'w', encoding='utf-8') as output_file:
        for personal_data in input_list:
            for key in personal_data.keys():
                if key == 'id' or \
                personal_data[key] == None:
                    continue
                for policy in policy_list:
                    if policy['字段名'] == key:
                        # format output
                        print('id:', personal_data['id'],
                            'result:',excute_func(policy['func_str'], personal_data[key]),
                            "字段名", policy['字段名'],
                            "数据产品名称", policy['数据产品名称'],
                            "规则名称", policy['规则名称'],
                            "规则内容", policy['规则内容'],
                            file=output_file)

def excute_func(func, input_param):
    return eval(func)(input_param)

if __name__ == '__main__':
    #policy_list = read_excel_data(r'C:\works\codes\Python\play_ground\银行新版经验规则集.xlsx')
    #to_analyze_personal_data_list = read_person_list(r'C:\works\codes\Python\play_ground\data_sample.xlsx')
    print(generate_policy_list_from_xls(excel_name=r'../policy_generator/银行新版经验规则集.xlsx'))
    #policy_soaking(to_analyze_personal_data_list, policy_list)