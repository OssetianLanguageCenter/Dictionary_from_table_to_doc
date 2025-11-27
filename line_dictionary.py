# класс, описывающий одну статью словаря, хранящуюся в строке таблицы
class Line_dictionary:
    def __init__(self, xls_sheet):
        self.sheet = xls_sheet
        self.examples = list()
        self.transc = list()
        self.rus = ""

    def read_line(self, row):
        line = Line_dictionary(self.sheet)
        #if self.sheet.cell(row, 0).value:
        try:
            line.rus = self.sheet.cell(row, 0).value
            line.reduct = self.sheet.cell(row, 1).value.strip()
            line.multy = self.sheet.cell(row, 3).value.strip()
            line.transc.clear()
            # если есть осетинский перевод
            if self.sheet.cell(row, 4).value.strip():
                line.transc.append((self.sheet.cell(row, 4).value.strip(), self.sheet.cell(row, 6).value.strip(), self.sheet.cell(row, 5).value.strip(), self.sheet.cell(row, 2).value.strip()))
            line.examples.clear()
            for i in range(7, 17, 2):
                if i + 1 < self.sheet.ncols:
                    rus_exp = self.sheet.cell(row, i).value.strip()
                    ose_exp = self.sheet.cell(row, i + 1).value.strip()
                    if rus_exp:
                        if rus_exp[-1] == ";":
                            rus_exp = rus_exp[:-1]
                        try:
                            if ose_exp[-1] == ";":
                                ose_exp = ose_exp[:-1]
                        except:
                            print(f"в строке {row} нет осетинского примера для существующего русского {rus_exp}!")
                        line.examples.append((rus_exp, ose_exp))
            print("rus ", line.rus)
            return line
        except:
            print(f"в строке {row} пусто")
            return  None
    #одно значение многозначного слова
class Polisemy:
    def __init__(self):
        self.multy = "" # номер значения
        self.transc = list() # список четверок слово - транскрипция - форма множественного числа или др - комментарий
        self.examples = list() # список пар пример - перевод

    def __str__(self):
        return  str(self.multy) + " !! " + str(self.transc) + " !! " + str(self.examples)

class Article:
    def __init__(self, sheet):
        self.rus = ""
        self.reduct = ""
        self.notice = ""
        self.polisemy = list()
        self.sheet = sheet

    def line_to_article(self, line):
        arcticle = Article(self.sheet)
        arcticle.rus = line.rus
        arcticle.reduct = line.reduct

        polisemy = Polisemy()
        polisemy.multy = line.multy
        polisemy.transc += line.transc
        polisemy.examples += line.examples
        arcticle.polisemy.append(polisemy)
        return arcticle

    def read_article(self, row):
        lines = Line_dictionary(self.sheet)
        line = lines.read_line(row)
        print("line", line.rus)
        arcticle = Article(self.sheet).line_to_article(line)
        line1 = lines.read_line(row + 1)
        print(line1)
        while line1 and not line1.rus:
            if line1.multy:
                polisemy = Polisemy()
                polisemy.multy = line1.multy
                if line1.transc: polisemy.transc += line1.transc
                if line1.examples: polisemy.examples += line1.examples
                arcticle.polisemy.append(polisemy)
            else:
                arcticle.polisemy[-1].transc += line1.transc
                if line1.examples: arcticle.polisemy[-1].examples += line1.examples
            row += 1
            line1 = lines.read_line(row + 1)
        return row, arcticle

    def __str__(self):
        temp = self.rus + " " + self.reduct + " " + self.notice
        for item in self.polisemy:
            temp += "\n" + str(item)
        temp += "\n_________________________\n"
        return temp
