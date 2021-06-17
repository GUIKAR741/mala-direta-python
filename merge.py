import csv
from mailmerge import MailMerge

# with MailMerge('base.docx') as document:
# print(document.get_merge_fields())
with open("resumos.csv") as m:
    reader = csv.reader(m)
    next(reader)
    for modalidade, titulo, resumo, palavras in reader:
        # print(modalidade, titulo, resumo, palavras)
        document = MailMerge('modelo.docx')
        print(document.get_merge_fields())
        document.merge(
            modalidade=modalidade,
            titulo=titulo,
            resumo=resumo,
            palavras=palavras
        )
        document.write(f'arq/{modalidade}.docx')
