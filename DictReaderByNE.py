import csv


class DictReaderByNE:
    def __init__(self, csvfile, *args, **kwargs):
        self.result = []
        self.data_type = ''
        with open(csvfile, *args, **kwargs) as csv_file:
            data_list = csv.reader(csv_file)
            title = None
            dict_add = {}
            for data in data_list:
                dict_temp = {}
                if len(data) == 1:
                    next_line = next(data_list)
                    if len(next_line) == 1:
                        dict_add.update({data[0].replace('[', '').replace(']', ''): next_line[0]})
                    else:
                        self.data_type = data[0].replace('[', '').replace(']', '')
                        title = next_line
                    continue
                if not data:
                    continue
                dict_temp.update(dict_add)
                dict_temp.update(dict(zip(title, data)))
                self.result.append(dict_temp)


if __name__ == '__main__':
    d = DictReaderByNE('Inventory_GAOI1_20210520_150300.csv')
    print(d.data_type)
    for a in d.result:
        print(a)
