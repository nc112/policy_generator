#数字转换为字母排序的表, 用于excel表格的转换
translate_table = {
        '0':'0',
        '1':'A',
        '2':'B',
        '3':'C',
        '4':'D',
        '5':'E',
        '6':'F',
        '7':'G',
        '8':'H',
        '9':'I',
        '10':'J',
        '11':'K',
        '12':'L',
        '13':'M',
        '14':'N',
        '15':'O',
        '16':'P',
        '17':'Q',
        '18':'R',
        '19':'S',
        '20':'T',
        '21':'U',
        '22':'V',
        '23':'W',
        '24':'X',
        '25':'Y',
        '26':'Z'
    }

def translate_numbers_to_words(target_column):
    assert target_column != None
    
    number_of_units_digit = str(target_column % 26)
    number_of_tens_digit = str(int(target_column / 26))
    translate_col = translate_table[number_of_tens_digit] + translate_table[number_of_units_digit]
    
    #remove 0. e.g. 0F -> F
    translate_col = translate_col.replace('0','')
    return translate_col