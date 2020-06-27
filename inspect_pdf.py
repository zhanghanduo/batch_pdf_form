import os
import pathlib
import pdfrw
import PySimpleGUI as sg

def run():
    temp_ = pdfrw.PdfReader('./template.pdf')
    print(temp_)
    print(temp_.Root)
    print(temp_.pages)
    # print(temp_.Root.Pages.Kids[0].Annots[0])
    # print(temp_.Root.Pages.Kids[0].Annots[1])
    # print(temp_.Root.Pages.Kids[0].Annots[2])
    # print(temp_.Root.Pages.Kids[0].Annots[3])
    # print(temp_.Root.Pages.Kids[0].Annots[4])


if __name__ == '__main__':
    merged_path = str(pathlib.Path().absolute()) + "\\" + os.path.basename('merged.pdf')
    os.makedirs(os.path.dirname(merged_path), exist_ok=True)
    t1 = pdfrw.PdfReader('./t1.pdf')
    t2 = pdfrw.PdfReader('./t2.pdf')
    writer_merge = pdfrw.PdfWriter()
    writer_merge.addpages(t1.pages)
    writer_merge.addpages(t2.pages)
    writer_merge.write(merged_path)


    # run()
    # for item in sg.list_of_look_and_feel_values():
    #     print(item)