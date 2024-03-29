# encoding=utf-8

import jieba

jieba.set_dictionary("dict.txt")
jieba.initialize()

ci_map = {}
sen_price = {}
sen_click_count_pc = {}
sen_click_count_mobile = {}
valid_count = 1
not_valid_char = ['\n']
total_word_count = 0


def clear_data():
    global ci_map
    global sen_price
    global sen_click_count_pc
    global sen_click_count_mobile
    ci_map = {}
    sen_cutted = {}
    sen_click_count_pc = {}
    sen_click_count_mobile = {}


def update_word_info(sen_cutted, sen):
    global ci_map
    for ci in sen_cutted:
        if ci in not_valid_char:
            continue
        d = ci_map.get(ci)
        if d:
            d['count'] = d['count'] + 1
        else:
            ci_map.setdefault(ci, {'count': 1})

        d = ci_map.get(ci)
        sen_list = d.get('sen_list', [])
        if sen not in sen_list:
            sen_list.append(sen)
        d['sen_list'] = sen_list


def handle_sen(sen):
    update_word_info(jieba.cut(sen, HMM=False), sen)


def anylizer():
    global ci_map
    global total_word_count
    ci_map_sorted = sorted(ci_map.items(), key=lambda kv: -kv[1]['count'])
    for k, v in ci_map_sorted:
        if v.get('count') >= valid_count:
            total_word_count = total_word_count + 1
            # print('%s %s' % (k, v))

    print('total_word_count=', total_word_count)


def get_sen_list_by_word(word):
    return ci_map.get(word, {}).get('sen_list', [])


def get_sen_price_by_sen(sen):
    return sen_price.get(sen, '-1')


def handle_sen_price_dict(sen, price):
    if sen not in sen_price.keys():
        sen_price.update({sen: price})


def get_sen_click_count_pc_by_sen(sen):
    return sen_click_count_pc.get(sen, '-1')


def handle_pc_sen_click_count(sen, click_count):
    if sen not in sen_click_count_pc.keys():
        sen_click_count_pc.update({sen: click_count})


def get_sen_click_count_mobile_by_sen(sen):
    return sen_click_count_mobile.get(sen, '-1')


def handle_mobile_sen_click_count(sen, click_count):
    if sen not in sen_click_count_mobile.keys():
        sen_click_count_mobile.update({sen: click_count})


def get_anylized_word_map():
    '''
    按照频率由高到低排序
    d = {word: {word_appear_count, ci_xing}
    return d
    '''
    ci_map_sorted = sorted(ci_map.items(), key=lambda kv: -kv[1]['count'])
    return ci_map_sorted
