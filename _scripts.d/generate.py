import xlrd
# import jinja2
import re



def replace_href(cell):
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        value = str(int(cell.value))
    else:
        value = str(cell.value)

    value = re.sub(r"\\href{(.*?)}{(.*?)}", r"[\2](\1)", value)
    value = re.sub(r"\\cite{(.*?)}", r"", value)
    value = re.sub(r"\\textbf{(.*?)}", r"<strong>\1<strong>", value)
    value = re.sub(r"\\textsc{(.*?)}", r'<font style="font-variant: small-caps">\1</font>', value)
    value = re.sub(r"\\dataindex{(.*?)}", r"\1", value)
    value = re.sub(r"\\rowcolor{(.*?)}", r"", value)
    return value
    # return str(cell.value)

class IdxCvt:
    def __init__(self, file):
        # TODO read to a dict
        self._data = dict()

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

    subsection_title = f"## {sheet.cell_value(0,0)}\n"
    content          = f"{sheet.cell_value(0,1)}\n"

    table_build = [subsection_title, content]

    table_build.append('|' + '|'.join(map(lambda cell: replace_href(cell), sheet.row(1))) + '|')
    table_build.append('|' + '|'.join(['-------'] * len(sheet.row(1))) + '|')
    i = 2
    while i < sheet.nrows:
        table_build.append('|' + idxcvt(sheet.cell_value(i,0)) + '|' + '|'.join(map(lambda cell: replace_href(cell), sheet.row(i)[1:])) + '|')
        i += 1

    return '\n'.join(table_build)




if __name__ == "__main__":
    import sys

    book = xlrd.open_workbook('Z:\\documents\\Medical dataset survey\\website\\heck-neck.xlsx')
    sheet = book.sheet_by_index(0)

    idxcvt = IdxCvt('')
    idxcvt._data['data:fastmri'] = "1"

    print(generate_table(sheet, idxcvt))