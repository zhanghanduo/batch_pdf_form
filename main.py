#!/usr/bin/env python3
import csv
import io
import os
import sys
import xlrd
import pdfrw
import PySimpleGUI as sg
from datetime import date
from subprocess import run, PIPE
if os.name == 'nt':
    import ctypes
    from ctypes import windll, wintypes
    from uuid import UUID
# from locale import atof, setlocale, LC_NUMERIC

index = 0
# 0 not initialized, 1 error, 2 file read, 3 converting, 4 interrupted, 5 finished
status = 0
max_row = 0
col_dict = {
    "l": 11, "k": 10, "j": 9, "i": 8, "h": 7, "g": 6, "f": 5, "e": 4, "d": 3, "c": 2}
col = 11
table_data = [[]]
header_list = []
sg.theme('LightBrown3')
date_ = date.today()
date_digit = date_.strftime('%d%m%y')
prefix_path = ''

class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", wintypes.DWORD),
        ("Data2", wintypes.WORD),
        ("Data3", wintypes.WORD),
        ("Data4", wintypes.BYTE * 8)
    ] 

    def __init__(self, uuid_):
        ctypes.Structure.__init__(self)
        self.Data1, self.Data2, self.Data3, self.Data4[0], self.Data4[1], rest = uuid_.fields
        for i in range(2, 8):
            self.Data4[i] = rest>>(8 - i - 1)*8 & 0xff


class UserHandle:
    current = wintypes.HANDLE(0)
    common  = wintypes.HANDLE(-1)


def get_path(folderid, user_handle=UserHandle.current):
    _CoTaskMemFree = windll.ole32.CoTaskMemFree
    _CoTaskMemFree.restype= None
    _CoTaskMemFree.argtypes = [ctypes.c_void_p]

    _SHGetKnownFolderPath = windll.shell32.SHGetKnownFolderPath
    _SHGetKnownFolderPath.argtypes = [
        ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_wchar_p)
    ] 

    fid = GUID(folderid) 
    pPath = ctypes.c_wchar_p()
    S_OK = 0
    if _SHGetKnownFolderPath(ctypes.byref(fid), 0, user_handle, ctypes.byref(pPath)) != S_OK:
        raise PathNotFoundException()
    path = pPath.value
    _CoTaskMemFree(pPath)
    return path


if os.name == 'nt':
    doc_id = UUID('{FDD39AD0-238F-46AF-ADB4-6C85480369C7}')
    doc_path = get_path(doc_id)
    prefix_path = doc_path + "\\" + 'filled'
else:
    doc_path = str(os.path.join(Path.home(), 'Documents'))
    prefix_path = doc_path + "/" + 'filled'


layout = [[sg.Text('输入日期 (如:190520): ', font=("Helvetica", 16)), 
            sg.Input(default_text=date_digit, font=("Helvetica", 16), key='-date-', enable_events=True, size=(6, 1)),
            sg.CalendarButton('选择日期', font=("Helvetica", 16), auto_size_button=True, target='-date-', format='%d%m%y', default_date_m_d_y=(6,19,2020), ),
            sg.StatusBar(text='默认今天', font=("Helvetica", 13), key='date_update', size=(12, 1))],
            [sg.Text('导入Excel/CSV 数据: ', font=("Helvetica", 16)), 
            sg.Input(key='-file-', enable_events=True, font=("Helvetica", 13), size=(30, 1)), sg.FileBrowse(font=("Helvetica", 16))],
            [sg.Text('输出文件夹名称', font=("Helvetica", 16)), 
            sg.Input(default_text=prefix_path, size=(30, 1), font=("Helvetica", 13), key='-output-', enable_events=True),
            sg.FolderBrowse(font=("Helvetica", 16))],
            [sg.Button('批量生成PDF', font=("Helvetica", 16)), sg.Button('Exit', font=("Helvetica", 16)),
            sg.StatusBar(text=' ', key='file_update', font=("Helvetica", 16), size=(12, 1), auto_size_text=True, pad=(10, 0)),
            sg.Button('打开生成文件夹', font=("Helvetica", 16), key='-view-', visible=False)],
            [sg.Text('金额所在列:', font=("Helvetica", 16)),
            sg.Combo(values=['L', 'K', 'J', 'I', 'H', 'G', 'F', 'E', 'D', 'C'], default_value='L', font=("Helvetica", 13), pad=(3, 3), key='-col-', enable_events=True),
            # sg.Input(default_text="L", font=("Helvetica", 13), key='-col-', enable_events=True, size=(2, 1)),
            sg.ProgressBar(max_value=10, orientation='h', size=(40, 22), key='progress', visible=False)]]

window = sg.Window('Cheque Excel to PDF Converting System', layout, location=(200, 140))
progress_bar = window['progress']


# def inspect_pdfs(args):
#     try:
#         with open(args.field_defs, "r") as f:
#             field_defs = json.load(f)
#     except OSError:
#         field_defs = {}
#     for filename in args.pdf_file:
#         field_defs[filename] = inspect_pdf_fields(filename)
#     with open(args.field_defs, "w") as f:
#         json.dump(field_defs, f, indent=4)
#     test_data = generate_test_data(args.pdf_file, field_defs)
#     fg = fill_forms(args.prefix, field_defs, test_data, True)
    # for filepath in fg:
    #     print(filepath)


def fill_pdfs(form_data, prefix):
    global status
    global prefix_path
    status = 3  #working
    fg = fill_forms_simple(prefix, form_data)
    # for filepath in fg:
    #     print(filepath)


def read_data(instream, datetime='today'):
    global status
    global max_row
    global header_list
    global table_data
    global col
    form_data = {}
    header_list = []
    table_data = [[]]
    max_row = 0
    status = 1
    if datetime == 'today':
        date_ = date.today()
        date_digit = date_.strftime('%d%m%y')
    elif str(datetime).isdigit():
        date_digit = str(datetime)

    if instream.endswith('.csv'):
        with open(instream, encoding='utf-8') as csvfile:
            reader_ = csv.reader(csvfile)
            next(reader_)
            header_list.extend([' No. ', '          Cust Name          ', ' 总额($) '])
            for row in reader_:
                if row and row[col] and row[1]:
                    max_row +=1
                    n_ = str(row[1]).split(" -")[0]
                    form_data[n_] = f = {}
                    list_ = []
                    list_.extend([max_row, n_, str(row[col])])
                    table_data.append(list_)
                    f["0"] = date_digit
                    f["1"] = n_
                    # a = atof(row[col])
                    a = float(str(row[col]).replace(',',''))
                    f["2"] = "$" + "{:.2f}".format(a)
                    f["4"] = None
                    f["3"] = "{:.2f}".format(a)
        status = 2
    elif instream.endswith('.xlsx') or instream.endswith('.xls'):
        wb = xlrd.open_workbook(instream)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0, 0)
        for i in range(sheet.nrows):
            if i == 0:
                header_list.extend([' No. ', '          Cust Name          ', ' 总额($) '])
            elif sheet.cell_value(i, 1) and sheet.cell_value(i, col):
                max_row +=1
                n_ = str(sheet.cell_value(i, 1)).split(" -")[0]
                v_ = sheet.cell_value(i, col)
                form_data[n_] = f = {}
                f["0"] = date_digit
                f["1"] = n_
                # a = atof(str(v_))
                a = float(str(v_).replace(',',''))
                # f["2"] = "$" + "{:,.2f}".format(a)
                f["2"] = "$" + "{:.2f}".format(a)
                f["4"] = None
                f["3"] = "{:.2f}".format(a)
                list_ = []
                list_.extend([max_row, n_, f["3"]])
                table_data.append(list_)
        status = 2
    else:
        status = 1
        sg.popup_error('Error reading file')
    return form_data


# def load_field_defs(defs_file):
#     with open(defs_file) as f:
#         return json.load(f)


# def inspect_pdf_fields(form_name):
#     cmd = ["pdftk", form_name, "dump_data_fields", "output", "-"]
#     p = run(cmd, stdout=PIPE, universal_newlines=True, check=True)
#     num = itertools.count()
#     fields = {}
#     for line in p.stdout.splitlines():
#         content = line.split(": ", 1)
#         if ["---"] == content:
#             fields[str(next(num))] = field_data = {}
#         elif 2 == len(content):
#             key = content[0][5:].lower()
#             if "stateoption" == key:
#                 field_data.setdefault(key, []).append(content[1])
#             else:
#                 field_data[key] = content[1]
#     return fields


def fill_forms(prefix, field_defs, data, flatten=True):
    global status
    global index
    progress_bar.update(visible=True)
    window.VisibilityChanged()
    window['file_update'].update('生成中~~~')
    for filename, formdata in data.items():
        if not formdata:
            continue
        # yield filename
        filepath = filename + '.pdf'
        output_path = make_path(prefix, filepath)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        fdf_str = generate_fdf(field_defs[".\\template.pdf"], formdata)
        fill_form(filepath, fdf_str, output_path, flatten)
        index += 1
        progress_bar.update_bar(index, max_row)
    status = 5
    index = 0


def fill_forms_simple(prefix, data):
    global status
    global index
    global output_path
    # template_pdf = pdfrw.PdfReader('./template.pdf')
    progress_bar.update(visible=True)
    window['file_update'].update('生成中~~~')
    for filename, formdata in data.items():
        if not formdata:
            continue
        # yield filename
        filepath = filename + '.pdf'
        output_path = make_path(prefix, filepath)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        temp_ = pdfrw.PdfReader('./template.pdf')
        temp_.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        for n, d in formdata.items():
            if d is not None:
                temp_.Root.Pages.Kids[0].Annots[int(n)].update(pdfrw.PdfDict(V=d))
        pdfrw.PdfWriter().write(output_path, temp_)
        index += 1
        progress_bar.update_bar(index, max_row)
    status = 5
    index = 0

'''
def generate_fdf(fields, data):
    fdf = io.StringIO()
    fdf.write(fdf_head)
    fdf.write("\n".join(fdf_fields(fields, data)))
    fdf.write(fdf_tail)
    return fdf.getvalue()


fdf_head = """%FDF-1.2
%âãÏÓ
1 0 obj 
<< /FDF 
<< /Fields [
"""

fdf_tail = """
] >> >>
endobj 
trailer
<< /Root 1 0 R >>
%%EOF
"""


def fdf_fields(fields, data):
    template = "<< /T ({field_name}) /V ({data}) >>"
    for n, d in data.items():
        field_def = fields.get(n)
        if field_def:
            field_name = field_def.get("name")
            if field_name:
                yield template.format(field_name=field_name, data=d)


def fill_form(input_path, fdf, output_path, flatten):
    cmd = ["pdftk", './template.pdf',
            "fill_form", "-",
            "output", output_path]
    if flatten:
        cmd.append("flatten")
    run(cmd, input=fdf.encode("utf-8"), check=True)


def generate_test_data(pdf_files, field_defs):
    data = {}
    for filepath in pdf_files:
        fields = field_defs.get(filepath, {})
        data[filepath] = d = {}
        for field_id, field_def in fields.items():
            if "Text" == field_def.get("type"):
                d[field_id] = field_id
    return data

'''


def make_path(prefix, path):
    return prefix + "\\" + os.path.basename(path)


def main():
    prefix = None
    date_ = 'today'
    global status
    global max_row
    global col
    global prefix_path
    form_data = {}
    table_exist = False

    while True:
        event, values = window.read()

        # date_time, file_path = values[0], values[1]
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == '-date-':
            date_ = values['-date-']
            window['date_update'].update('日期已输入')
        if event == '-file-':
            window['file_update'].update('文件地址已输入')
            form_data = read_data(values['-file-'], date_)
            if table_exist:
                window['file_update'].update('数据已经更新')
                window['-table-'].update(values=table_data, num_rows=min(len(table_data), 20))
            else:
                table_exist = True
                window['file_update'].update('数据已经导入')
                window.extend_layout(window, [[sg.Table(values=table_data, headings=header_list, max_col_width=14, 
                auto_size_columns=True, justification='left', alternating_row_color='lightyellow', header_text_color='teal',
                font=("Helvetica", 13), key='-table-', num_rows=min(len(table_data), 20))]])

        if event == '-output-':
            prefix = values['-output-']
        if event == '批量生成PDF':
            if status == 2 and form_data:
                if prefix is None:
                    fill_pdfs(form_data, prefix_path)
                else:
                    fill_pdfs(form_data, str(prefix))
                    prefix_path = str(prefix)
        if event == '-col-':
            if values['-col-']:
                input = values['-col-'].lower()
                col = col_dict[input]
                form_data = read_data(values['-file-'], date_)
                window['file_update'].update('数据已经更新')
                if table_exist:
                    window['file_update'].update('数据已经更新')
                    window['-table-'].update(values=table_data, num_rows=min(len(table_data), 20))
                else:
                    table_exist = True
                    window['file_update'].update('数据已经导入')
                    window.extend_layout(window, [[sg.Table(values=table_data, headings=header_list, max_col_width=100, 
                    def_col_width=80, auto_size_columns=False, justification='left', alternating_row_color='lightblue', 
                    header_text_color='blue', font=("Helvetica", 13), key='-table-', num_rows=min(len(table_data), 20))]])
        if event == '-view-':
            if prefix_path:
                os.startfile(prefix_path)
                print(prefix_path)

        # print(event)
        if status == 5:
            window['file_update'].update('任务完成!')
            window['-view-'].update(visible=True)
            max_row = 0
            status = 0

        elif status == 1:
            window['file_update'].update('输入的文件格式不对!')
            max_row = 0
            status = 0

    window.close()


if __name__ == "__main__":
    main()
