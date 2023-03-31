from configparser import ConfigParser
import datetime

_DEFAULT_CONFIG = '''[vars]
cadfolder_path = C:/
blanks_path = ./Бланки
pp_header = PP_BLANK
position_marker = Поз.
quantity_marker = Кол-во
number_marker = Номер
spec_sort_1 = Обозначение
spec_sort_2 = Наименование
mergetocolumn = 9
author = 
other = 

[MASKS]
productname_mask = <PRODUCTNAME>
author_mask = <AUTHOR>
docsfolder_mask = <DOCSFOLDER>
dateofissue_mask = <DATEOFISSUE>
other_mask = <OTHER>
'''


def touch_config(force_create=False):
    configur = ConfigParser()
    if force_create:
        try:
            with open('config.ini', 'w') as conf:
                conf.write(_DEFAULT_CONFIG)
            return touch_config()
        except:
            raise RuntimeError
    else:
        try:
            masks = dict()
            configur.read("config.ini")
            for key, value in configur.items('MASKS'):
                masks.update({key.upper(): value})
            masks_values = {
                masks['DATEOFISSUE_MASK']: str(datetime.date.today()),
                masks['AUTHOR_MASK']: configur.get('vars', "author"),
                masks['OTHER_MASK']: configur.get('vars', 'other')
            }
            out_params = dict()
            for key, value in configur.items('vars'):
                out_params.update({str(key).upper(): value})
            out_params.update({'MASKS': masks})
            out_params.update({'MASKS_VALUES': masks_values})
            return out_params
        except Exception:
            return touch_config(force_create=True)


def config_save(params: dict):
    vars_dict = dict()
    out_dict = dict()
    for key, value in params.items():
        if not isinstance(value, dict):
            vars_dict.update({key: value})
        elif key == 'MASKS_VALUES':
            pass
        else:
            out_dict.update({key: value})
    out_dict.update({'vars': vars_dict})
    config = ConfigParser()
    config.read_dict(out_dict)
    with open('config.ini', 'w') as f:
        config.write(f)
