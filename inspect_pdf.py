import pdfrw
import PySimpleGUI as sg

def run():
    temp_ = pdfrw.PdfReader('./template.pdf')
    print(temp_.Root.Pages.Kids[0].Annots[0])
    print(temp_.Root.Pages.Kids[0].Annots[1])
    print(temp_.Root.Pages.Kids[0].Annots[2])
    print(temp_.Root.Pages.Kids[0].Annots[3])
    print(temp_.Root.Pages.Kids[0].Annots[4])


if __name__ == '__main__':
    # run()
    for item in sg.list_of_look_and_feel_values():
        print(item)