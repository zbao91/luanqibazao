# encoding: utf-8
import sys
import os
parent_dir = '/'.join(os.path.abspath(__file__).split('/')[:-2]) # ref: https://stackoverflow.com/questions/22955684/how-to-import-py-file-from-another-directory
sys.path.insert(0, parent_dir)

from common.excel import ExlProcess

if __name__ == "__main__":
    dir_path = "/Users/zhiqibao/Desktop/Life/郦城公馆电费/20200807"
    files = os.listdir(dir_path)
    excel_obj = ExlProcess()
    overall_data = {}
    final_data = []
    for file in files:
        if "总结" in file:
            continue
        file_path = os.path.join(dir_path, file)
        data = excel_obj.extract_data(file_path)
        for i in data["消费记录"]:
            date = i.get("账单日期").split(" ")[0]
            fee = float(i.get("支付金额(元)"))
            if not date in overall_data:
                overall_data[date] = fee
            else:
                overall_data[date] += fee
    for key, value in overall_data.items():
        tmp_dict = {
            "日期": key,
            "电费": value
        }
        final_data.append(tmp_dict)
    final_data = sorted(final_data, key=lambda x:x["日期"])

    excel_obj.write_data(os.path.join(dir_path, "电费总结.xlsx"), 0, final_data)



