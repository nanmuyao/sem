# encoding=utf-8
import glob
import csv
import jieba

from collections import OrderedDict

from pip._vendor import chardet

jieba.set_dictionary("dict.txt")
jieba.initialize()


def handle():
    titles, row_dict = load_data()
    keywords_dict = get_jieba_cut_keywords_count_dict(row_dict)
    write_to_csv(titles, row_dict, keywords_dict)


def get_vsvs():
    csvs = []
    path = "*.csv"
    for fname in glob.glob(path):
        csvs.append(fname)
    return csvs


def load_data():
    csvs = get_vsvs()
    for csv_name in csvs:
        with open(csv_name, 'rb') as fin:
            encoding_type = chardet.detect(fin.read(70))['encoding']

            # reader = csv.reader(fin, delimiter=' ', quotechar='|', )
            # for row in reader:
            #     print(row)
        row_dict = OrderedDict()
        titles = []
        with open(csv_name, newline='', encoding=encoding_type) as f:
            reader = csv.reader(f)
            for row in reader:
                title_row = row[0]
                titles = title_row.split('\t')
                break
            print('titles=', titles)

            for row in reader:
                print('row==', row)
                title_row = row[0]
                # row = title_row.rsplit()
                print('split row ', row)
                row = title_row.split('\t')
                print('22split row ', row)

                d = {}
                for index, value in enumerate(row):
                    print(index)
                    d[titles[index]] = value
                row_dict[d.get(titles[0])] = d

            print(row_dict)
            return titles, row_dict


def get_jieba_cut_keywords_count_dict(row_dict):
    # 把关键词切分，并且计算出关键词出现的次数
    keywords_dict = {}
    for keyword, search_content in row_dict.items():
        # update_word_info(jieba.cut(sen, HMM=False), sen)
        keys = jieba.cut(keyword, HMM=False)
        for key in keys:
            print('key==', key)

        if keyword in keywords_dict.keys():
            keywords_dict[keyword] = keywords_dict[keyword] + 1
        else:
            keywords_dict[keyword] = 1

    return keywords_dict


def write_to_csv(titles, row_dict, keywords_dict):
    with open('export.csvabc', 'w', newline='') as csvfile:
        _titles = titles.copy()
        del _titles[0]

        group_dict = {}
        group_index = 0
        export_datas = []
        fieldnames = []
        for keyword, value in keywords_dict.items():
            fieldnames.append('%s-%s' % (keyword, value))
            for title in _titles:
                _title = '%s-%s' % (title, group_index)
                fieldnames.append(_title)

            for k, v in row_dict.items():
                if keyword in k:
                    d = {}
                    for field_name in fieldnames:
                        v_key = field_name.split('-')[0]
                        d[field_name] = v.get(v_key)

                    export_datas.append(d)

            group_index += 1

        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerow(export_datas)


def export():
    pass


if __name__ == '__main__':
    handle()
