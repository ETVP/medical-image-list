import xlrd
# import jinja2
import re
import sys



def replace_latex(cell, format = None, ignore_bf = True):
    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return ""
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        value = str(int(cell.value))
    else:
        value = str(cell.value)

    value = re.sub(r"\\href{(.*?)}{(.*?)}", r"[\2](\1)", value).strip()
    value = re.sub(r"\\cite{(.*?)}", r"", value).strip()
    value = re.sub(r"\\textbf{(.*?)}", r"\1" if ignore_bf else r"**\1**", value).strip()
    value = re.sub(r"\\textsc{(.*?)}", r'\1', value).strip()
    value = re.sub(r"\\dataindex{(.*?)}", r"\1", value).strip()
    value = re.sub(r"\\rowcolor{(.*?)}", r"", value).strip()
    value = re.sub(r"\\&", r"&", value)
    value = format.format(value = value) if format is not None else value
    return value
    # return str(cell.value)

class IdxCvt:
    def __init__(self, file):
        self._data = dict()
        for line in open(file, "r", encoding="utf-8"):
            token  = line.split(',')
            self._data[token[0].strip()] = token[1].strip()


    def __call__(self, key):
        key = re.sub(r"\\dataindex{(.*?)}", r"\1", key)
        return self._data.get(key, key)

def idxId(idx):
    return re.sub(r"\\dataindex{(.*?)}", r"\1", idx)


def generate_table(sheet, idxcvt = idxId):
    """
    A1 是小节的名字，A2 是表头
    2 开始为 “表本身”。
    B
    """

    subsection_title = f"\n## {replace_latex(sheet.cell(0,0), ignore_bf = False)}\n"
    content          = f"{replace_latex(sheet.cell(0,1), ignore_bf = False)}\n"

    table_build = [subsection_title, content]

    table_build.append('|' + '|'.join(map(lambda cell: replace_latex(cell, '<span style="font-variant: small-caps; font-weight:bold;"> {value} </span>'), sheet.row(1))) + '|')
    table_build.append('|' + '|'.join(['-------'] * len(sheet.row(1))) + '|')
    i = 2
    while i < sheet.nrows:
        # print('Row', i, file=sys.stderr)
        table_build.append('|' + idxcvt(sheet.cell_value(i,0)) + '|' + '|'.join(map(lambda cell: replace_latex(cell), sheet.row(i)[1:])) + '|')
        i += 1

    return '\n'.join(table_build)




if __name__ == "__main__":

    idxcvt = IdxCvt('B:\\etvp\\Paper-survey-medical-dataset\\main.dataindex')
    # book = xlrd.open_workbook('Z:\\documents\\Medical dataset survey\\website\\head-neck.xlsx')
    # book = xlrd.open_workbook('Z:\\documents\\Medical dataset survey\\website\\chest-abdomen.xlsx')
    # book = xlrd.open_workbook('Z:\\documents\\Medical dataset survey\\website\\pathology-blood.xlsx')
    book = xlrd.open_workbook('Z:\\documents\\Medical dataset survey\\website\\others.xlsx')

    for i in range(book.nsheets):
        # print('Sheet', i, file=sys.stderr)
        sheet = book.sheet_by_index(i)
        table = generate_table(sheet, idxcvt)
        # import pdb; pdb.set_trace()
        print(table)