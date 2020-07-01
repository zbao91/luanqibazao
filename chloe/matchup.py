# encoding: utf-8
import sys
import os
parent_dir = '/'.join(os.path.abspath(__file__).split('/')[:-2]) # ref: https://stackoverflow.com/questions/22955684/how-to-import-py-file-from-another-directory
sys.path.insert(0, parent_dir)

from common.excel import ExlProcess

if __name__ == '__main__':
    file_path = '/Users/zhiqibao/Desktop/Project/files/chloe_files.xlsx'
    xl_obj = ExlProcess()
    data = xl_obj.extract_data(file_path, way='index', sheet_name=[0])
    mem_dict = {}
    for sheet in data:
        for row in data[sheet]:
            _id = row.get('会员号')
            counter = row.get('柜台名')
            times = row.get('领取件数')
            if counter == 'EC':
                counter = 'online'
            else:
                counter = 'offline'
            if not _id in mem_dict:
                mem_dict[_id] = {'data': [row], 'online': 0, 'offline': 0}
                mem_dict[_id][counter] += times
            else:
                mem_dict[_id][counter] += times
                mem_dict[_id]['data'].append(row)

    sheet_1 = [] # online and offline, offline once
    sheet_2 = [] # online and offline, offline multi
    sheet_3 = [] # online only, online multi
    sheet_4 = [] # offline only, offline multi

    for _id in mem_dict:
        tmp_data = mem_dict[_id].get('data')
        online_count = mem_dict[_id]['online']
        offline_count = mem_dict[_id]['offline']
        if online_count > 0 and offline_count == 1:
            sheet_1 += tmp_data
        elif online_count > 0 and offline_count > 1:
            sheet_2 += tmp_data
        elif offline_count == 0 and online_count > 1:
            sheet_3 += tmp_data
        elif online_count == 0 and offline_count > 1:
            sheet_4 += tmp_data
        else:
            pass

    xl_obj.write_data(file_path, 1, sheet_1)
    xl_obj.write_data(file_path, 2, sheet_2)
    xl_obj.write_data(file_path, 3, sheet_3)
    xl_obj.write_data(file_path, 4, sheet_4)


