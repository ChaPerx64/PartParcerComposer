import os


class CDPath:
    def __init__(self, path: os.PathLike):
        self.path = str(path)


def is_gc_directory(path: str):
    pathway_list = path.split('/')
    if pathway_list[len(pathway_list) - 1] == 'GrabCAD':
        return True
    else:
        return False


def is_in_gc(path: str):
    if 'GrabCAD' in path:
        return True
    else:
        return False


def is_project_folder(path: str):
    pathway_list = path.split('/')
    if pathway_list[len(pathway_list) - 2] == 'GrabCAD':
        return True
    else:
        return False


def get_project_pathway(path: str):
    if is_in_gc(path) and not is_gc_directory(path):
        pathway_list = path.split('/')
        while pathway_list[len(pathway_list)-2] != 'GrabCAD':
            # print(pathway_list[len(pathway_list)-2])
            pathway_list.pop(len(pathway_list)-1)
            # pass
        path_out = '/'.join(pathway_list)
        return path_out
    else:
        raise NotADirectoryError('Этот путь не ведет в ПОДпапку GrabCAD!')

# pth = 'D:/Workfolder/Workfolder_CAD/Photomechanics/GrabCAD/Familia/PMA.RCK.00.000/CAD/PMA.RCK.00.000.SLDASM'
# print(get_project_pathway(pth))