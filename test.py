import io

import pdfrw

def run():
    template = pdfrw.PdfReader('./template.pdf')

    # template.Root.Pages.Kids[0].Annots[0].update(pdfrw.PdfDict(V='(test)'))
    template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
    template.Root.Pages.Kids[0].Annots[0].update(pdfrw.PdfDict(V='302620'))
    template.Root.Pages.Kids[0].Annots[1].update(pdfrw.PdfDict(V='Random1 Name Pte Ltd)'))
    pdfrw.PdfWriter().write('test.pdf', template)


# def get_overlay_canvas() -> io.BytesIO:
#     data = io.BytesIO()
#     pdf = canvas.Canvas(data)
#     pdf.drawString(x=33, y=550, text='Willis')
#     pdf.drawString(x=148, y=550, text='John')
#     pdf.save()
#     data.seek(0)
#     return data


# def merge(overlay_canvas: io.BytesIO, template_path: str) -> io.BytesIO:
#     template_pdf = pdfrw.PdfReader(template_path)
#     overlay_pdf = pdfrw.PdfReader(overlay_canvas)
#     for page, data in zip(template_pdf.pages, overlay_pdf.pages):
#         overlay = pdfrw.PageMerge().add(data)[0]
#         pdfrw.PageMerge(page).add(overlay).render()
#     form = io.BytesIO()
#     pdfrw.PdfWriter().write(form, template_pdf)
#     form.seek(0)
#     return form


def save(form: io.BytesIO, filename: str):
    with open(filename, 'wb') as f:
        f.write(form.read())

if __name__ == '__main__':
    run()