class EntryMatcher:
    def __init__(self, filters: dict, partslist: list[dict], params: dict):
        self.errors = []
        self.partslist = partslist
        self.filters = filters
        self.params = params
        self.check_spec()

    def check_spec(self):
        entry = self.partslist[0]
        for key in self.filters.keys():
            if not key in entry.keys():
                self.errors.append('В исходной спеке отстутствует заголовок  "' + key + '"')
        if self.errors:
            return self.errors
        else:
            return None

    def get_matches(self):
        ls_out = list()
        if self.errors:
            ls_out.append('Произошли следующие ошибки:')
            for item in self.errors:
                ls_out.append(item)
        else:
            for entry in self.partslist:
                flag1 = True
                entry_out = dict()
                for key, filter_ls in self.filters.items():
                    for subfilter in filter_ls:
                        if str(subfilter).lower() in str(entry.get(key)).lower():
                            entry_out.update({key: entry.get(key)})
                        else:
                            flag1 = False
                            break
                if flag1:
                    ls_out.append(entry_out)
        return ls_out