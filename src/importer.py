# -*- encode: utf-8 -*-
import xlrd
import csv
import datetime
import time

class BankMovement(object):
    def __init__(self, date, concept, price, balance):
        self.date = date
        self.concept = concept
        self.price = price
        self.balance = balance

    def date_formated(self):
        return self.date.strftime('%d/%m/%y')
    
    def export(self, field_delimiter):
        return field_delimiter.join([self.date_formated(),
                                     (self.concept),
                                     str(self.price),
                                     str(self.balance)])

    def toCsv(self):
        return self.export('";"')

    def toHTML(self):
        return u'{0}\t{1}\t{2}\t{3}\n'.format(self.date_formated(),
                                              self.concept,
                                              self.price,
                                              self.balance)
                        
        
class BankImporter(object):

    def __init__(self, file):
        self._file = file
        self._movements = []
        self._parse_file()

    def _parse_file(self):
        pass

    def _float(self, number ):
        return float(number)

    def __iter__(self):
        return self._movements.__iter__()

    def __len__(self):
        return len(self._movements)
        


class INGBankImporter(BankImporter):

    BEGINING = (4,0)
    DATE_COLUMN_IDX = 0
    CONCEPT_IDX = 1
    PRICE_IDX = 2
    ACCOUNT_BALANCE_IDX = 3

    def _parse_file(self):

        workbook = xlrd.open_workbook(self._file)
        sheet = workbook.sheet_by_index(0)

        self._movements = [self._parse_row(sheet.row(row_id)) for row_id in xrange(INGBankImporter.BEGINING[0], sheet.nrows)]
        self._movements.sort(lambda x,y: 1 if x.date > y.date else -1)

        return self._movements
        

    def _parse_row(self, row):
        '''
        '''
        cell = row[INGBankImporter.DATE_COLUMN_IDX]
        
        # date = datetime.datetime(*xlrd.xldate_as_tuple(cell.value, 0))
        date = datetime.datetime.strptime(cell.value, '%d/%m/%Y')
        
        concept = row[INGBankImporter.CONCEPT_IDX].value
        
        item = BankMovement(date,
                            concept[0:concept.find('(')],
                            self._float(row[INGBankImporter.PRICE_IDX].value.replace(',', '.')),
                            self._float(row[INGBankImporter.ACCOUNT_BALANCE_IDX].value.replace('.', '').replace(',', '.')))
        return item
                        
    def toHTMLTable(self):
        '''
        '''
        html =''
        for move in self._movements:
            html += u'{0}\n'.format(move.toHTMLTable())

        print html


class OficinaDirectaBankImporter(BankImporter):
    DATE_COLUMN_IDX = 0
    CONCEPT_IDX = 1
    PRICE_IDX = 3 
    ACCOUNT_BALANCE_IDX = 4

    def _parse_file(self):
        reader = csv.reader(open(self._file, 'r'), delimiter="\t")
        for row in reader:
            movement = self._parse_row(row)
            if movement:
                self._movements.append(movement)
        
    def _parse_row(self, row):
        # Ignore blank lists
        if len(row) > 4:
            # Ignore header
            date = row[OficinaDirectaBankImporter.DATE_COLUMN_IDX].strip()
            if not date.upper().startswith('FECHA'):
                date = time.strptime(date, r'%d/%m/%Y')
                date = datetime.date(date.tm_year, date.tm_mon, date.tm_mday)
                return BankMovement(date,
                                    row[OficinaDirectaBankImporter.CONCEPT_IDX].strip(),
                                    self._float(row[OficinaDirectaBankImporter.PRICE_IDX]),
                                    self._float(row[OficinaDirectaBankImporter.ACCOUNT_BALANCE_IDX]))
        return None

    def _float(self, number):
        return float(number.replace(',',''))


def run(fd):
    ing = INGBankImporter(fd)  # OficinaDirectaBankImporter(fd) #
    
    for i in ing:
        print u"{}".format(i.toCsv().encode('ascii', 'ignore').decode('utf-8'))

if __name__ == '__main__':
    import sys
    run(sys.argv[1])
    