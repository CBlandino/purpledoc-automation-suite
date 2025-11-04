from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfString, PdfObject
from typing import Dict

def fill_pdf(input_pdf_path: str, output_pdf_path: str, data_dict: Dict[str, str]):
    template_pdf = PdfReader(input_pdf_path)
    annotations = template_pdf.pages[0].get('/Annots', [])
    if annotations:
        for annotation in annotations:
            if annotation['/Subtype'] == '/Widget' and annotation.get('/T'):
                key = annotation['/T'][1:-1]
                if key in data_dict:
                    value = data_dict.get(key) or ''
                    annotation.update(PdfDict(V=PdfString.encode(value)))
                    annotation.update(PdfDict(AS=PdfName('Yes')))
        if template_pdf.Root.AcroForm:
            template_pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject('true')))
        else:
            template_pdf.Root.update(PdfDict(AcroForm=PdfDict(NeedAppearances=PdfObject('true'))))
    PdfWriter().write(output_pdf_path, template_pdf)
