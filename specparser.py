# from openpyxl import load_workbook
import operator

from aux_code import load_first_sheet
# SPECS = 'specs'

def p(what='foo'):
    print(what)


class SpecParser:
    def __init__(self, xl_path, params: dict):
        self.path = xl_path
        self.sheet = self._load_first_sheet()
        self.pos_key = params.get('POSITION_MARKER')
        self.qty_key = params.get('QUANTITY_MARKER')
        self.key1 = params.get('KEY_1')
        self.key2 = params.get('KEY_2')
        self.specs_key = 'specs'

    def _load_first_sheet(self, remove_enters=True):
        return load_first_sheet(self.path)

    # Метод, получающий заголовки из листа
    def _get_spec_headers(self):
        headers = []
        for row in self.sheet.rows:
            for cell in row:
                if cell.value:
                    headers.append(str(cell.value))
                else:
                    headers.append('')
            break
        return headers

    def _get_spec_rows(self):
        rows = []
        counter = 1
        for row in self.sheet.rows:
            line = []
            if counter > 1:
                for cell in row:
                    if cell.value:
                        line.append(str(cell.value))
                    else:
                        line.append(None)
                rows.append(line)
            counter += 1
        return rows

    def get_entries_raw(self):
        entries = list()
        headers = self._get_spec_headers()
        rows = self._get_spec_rows()
        for row in rows:
            entry = dict()
            for header, value in zip(headers, row):
                entry.update({header: value})
            entries.append(entry)
        return entries

    def entry_to_dict(self, entry: dict):
        out_dict = dict()
        specs_dict = dict()
        for key, value in entry.items():
            if key == self.pos_key:
                out_dict.update({key: value})
            elif key == self.qty_key:
                out_dict.update({key: value})
            else:
                specs_dict.update({key: value})
        out_dict.update({self.specs_key: specs_dict})
        return out_dict

    def get_entries(self):
        entries = list()
        for entry in self.get_entries_raw():
            entries.append(self.entry_to_dict(entry))
        return entries

    def get_counted(self):
        prev_depth = 0
        prev_qty = 0
        mplicators_ls = [1]
        counted_list = list()
        for entry in self.get_entries():
            qty = int(entry[self.qty_key])
            depth = str(entry[self.pos_key]).count('.')
            if prev_depth < depth:
                mplicators_ls.append(prev_qty)
            if prev_depth > depth:
                mplicators_ls.pop(len(mplicators_ls)-1)
            mplicator = 1
            for a in mplicators_ls:
                mplicator = mplicator * a
            entry[self.qty_key] = str(qty * mplicator)
            counted_list.append(entry)
            prev_qty = qty
            prev_depth = depth
        return counted_list

    def get_flatlist(self):
        flat_list = list()
        for entry in self.get_counted():
            counted = False
            for fl_entry in flat_list:
                if fl_entry[self.specs_key] == entry[self.specs_key]:
                    fl_entry[self.qty_key] = str(int(entry[self.qty_key]) + int(fl_entry[self.qty_key]))
                    counted = True
                    break
            if not counted:
                flat_list.append(entry)
        return flat_list

    def dict_to_entry(self, dict_entry:dict):
        for key, value in dict_entry[self.specs_key].items():
            dict_entry.update({key: value})
        dict_entry.pop(self.specs_key)
        return dict_entry

    def get_flat_unformatted(self):
        out_list = list()
        for entry in self.get_flatlist():
            out_list.append(self.dict_to_entry(entry))
        out_list.sort(key=self._sort_weight)
        # out_list.sort(key=operator.itemgetter('Обозначение', 'Наименование'))
        return out_list

    def _sort_weight(self, line: dict):
        out_str = ''
        for key in (self.key1, self.key2):
            if line.get(key):
                out_str += str(line.get(key)).lower()
        return out_str




# for tests
if __name__ == '__main__':
    XL_PATH = './Пример спеки/Familia Rack.xlsx'
    params_now = {'POSITION_MARKER': 'Поз.', 'QUANTITY_MARKER': 'Кол-во'}
    parser = SpecParser(XL_PATH, params_now)
    count = 1
    for item in parser.get_flat_unformatted():
        print('\nEntry #' + str(count))
        print(item)
        count += 1

