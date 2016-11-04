import xlsxwriter
import win32com.client
import re
import os
import pandas as pd
import json

new_delim = "_$_"

def get_levels_dict(path):
    data = []
    with open(path) as f:
        data = json.load(f)
    return data

def get_levels_vars(levels, sav_vars):
    levels_vars = []
    i = 0
    for level in levels:
        levels_vars.append({ 'name' : level['name'], 'data': [] })
        levels_vars[i]['data'].append(*level['fixed_vars'])
        for entry in level['entries']:
            for q in entry['question_begining']:
                for var in sav_vars[1]:
                    if re.search('^'+ q + entry['delimeters'][0][0], var):
                        levels_vars[i]['data'].append(var)
        i += 1
    return levels_vars

def save_levels_to_excel(levels_vars, cur_path, spss):
    for level in levels_vars:
        with open(cur_path + '\\' + level['name'] + '.sps', 'w') as file:
            file.write("SAVE TRANSLATE OUTFILE='" + cur_path + '\\' + level['name'] + '.xlsx\'' +
                        """
                            /TYPE=XLS
                            /VERSION=12
                            /MAP
                            /REPLACE
                            /FIELDNAMES
                            /CELLS=VALUES
                            /KEEP=
                        """)
            for var in level['data']:
                file.write(var + '\n')
            file.write('.\nEXECUTE.\n')
        syntax = spss.OpenSyntaxDoc(cur_path + '\\' + level['name'] + '.sps')
        syntax.run()
        while spss.IsBusy():
            pass

def stack_dataset(level, data, stack):
    multiIndex = []
    for var in data.columns:
        for entry in level['entries']:
            for q in entry['question_begining']:
                if re.search('^'+q, var):
                    temp = []
                    temp_str = var
                    for spl in entry['delimeters'][stack]:
                        temp.append(temp_str.split(spl,1)[0])
                        temp_str = temp_str.split(spl,1)[1]
                    temp.append(temp_str)
                    while len(temp) < level['dept']:
                        temp.append(new_delim)
                    multiIndex.append(tuple(temp))
    data.columns = pd.MultiIndex.from_tuples(multiIndex)
    data.columns = data.columns.swaplevel(1,len(data.columns.levels)-1)
    data = data.stack()
    data.index.names = ([ *data.index.names[:-1], level['stack_name'][stack] ])
    new_cols = []
    for x in data.columns:
        temp = new_delim
        temp = temp.join(x)
        temp = temp.replace(new_delim + new_delim,"")
        new_cols.append(temp)
    data.columns = new_cols
    return data
    
def save_restructured_data_to_excel(level, cur_path, data):
    writer = pd.ExcelWriter(cur_path + '\\' + level['name'] + "_stacked.xlsx",
                                engine='xlsxwriter')
    data.to_excel(writer, sheet_name="Sheet1", merge_cells=False)
    writer.save()

def get_excel_dataset(level, cur_path):
    return pd.read_excel(cur_path + '\\' + level['name'] + '.xlsx', 
                             index_col=[x for x in range(0,len(level['fixed_vars']))])

def save_restructured_data_to_spss(level, cur_path, data, spss):       
    command = """GET DATA /TYPE=XLSX 
          /FILE='""" + cur_path + '\\' + level['name'] + """_stacked.xlsx'
          /SHEET=name 'Sheet1' 
          /CELLRANGE=full 
          /READNAMES=on 
          .
          EXECUTE.
        """ + 'DATASET NAME ' + level['name'] + '_ds.'
    spss.ExecuteCommands(command, True)
    while spss.IsBusy():
        pass
    target_doc = spss.Documents.GetDataDoc(spss.Documents.DataDocCount - 1)
    target_doc.visible = True
    spss.ExecuteCommands("DATASET ACTIVATE " + level['name'] + '_ds.', True)
    syntax = spss.OpenSyntaxDoc(cur_path + '\\' + level['name'] + '_labels.sps')
    syntax.run()
    target_doc.SaveAs(cur_path + '\\' + level['name'] + "_stacked.sav")

def produce_labels_syntax(level, cur_path, dataset, levels_vars, sav):
    with open(cur_path + '\\' + level['name'] + '_labels.sps', 'w') as f:
        for var in [*dataset.columns, *dataset.index.names]:
            for ori_var in range(0,levels_vars[0]):
                if levels_vars[1][ori_var] == var or re.search('^'+ var.split(new_delim)[0], levels_vars[1][ori_var]):
                    f.write("VARIABLE LABELS " + var + " '" + levels_vars[2][ori_var] + "'.\n")       
                    label = sav.GetVariableValueLabels(ori_var)
                    if len(label) > 0:
                        f.write("VALUE LABELS " + var + "\n" )
                        for i in range(0,len(label[1])):
                            f.write(str(label[1][i]) + " '" + str(label[2][i]) + "'\n")
                        f.write(".\n")
                    break

def main():
    spss = win32com.client.Dispatch("SPSS.Application16")
    sav_file = "C:\\Users\\plamen.tarkalanov\\Documents\\Work\\Amazon\\Amazon\\TEST\\92173773_(54881)_data_2016-02-26-00h28m.sav"
    sav = spss.OpenDataDoc(sav_file)

    cur_path = os.getcwd()
    levels = get_levels_dict("C:\\Users\\plamen.tarkalanov\\Documents\\Work\\Amazon\\Amazon\\TEST\\test.json")
    var_info = sav.GetVariableInfo()

    levels_vars = get_levels_vars(levels, var_info)
    save_levels_to_excel(levels_vars, cur_path, spss)

    for level in levels:
        dataset = get_excel_dataset(level, cur_path)
        for i in range(0, len(level['stack_name'])):
            dataset = stack_dataset(level, dataset, i)
        save_restructured_data_to_excel(level, cur_path, dataset)
        produce_labels_syntax(level, cur_path, dataset, var_info, sav)
        save_restructured_data_to_spss(level, cur_path, dataset, spss)
    
main()

