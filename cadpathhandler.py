import os


class CADPathHandler:
    def __init__(self, cadfolderpath: str):
        self._cadfolderpath = cadfolderpath
        self._cadfoldername = cadfolderpath.split('/')[-1]

    def is_in_cadfolder(self, path: str):
        if self._cadfolderpath in path:
            return True
        else:
            return None

    def folder_level(self, path: str):
        pathway_list = path.split('/')
        if self.is_in_cadfolder(path):
            return len(pathway_list) - pathway_list.index(self._cadfoldername) - 1
        else:
            return None

    def is_gc_directory(self, path: str):
        if self.folder_level(path) == 0:
            return True
        else:
            return None

    def is_project_folder(self, path: str):
        if self.folder_level(path) == 1:
            return True
        else:
            return None

    def get_project_pathway(self, path: str):
        if self.is_in_cadfolder(path) and not self.is_gc_directory(path):
            pathway_list = path.split('/')
            while pathway_list[-2] != 'GrabCAD':
                pathway_list.pop(-1)
            path_out = '/'.join(pathway_list)
            return path_out
        else:
            return None

    def strip_to_cadfolder(self, path):
        if self.is_in_cadfolder(path):
            path = path.replace(self._cadfolderpath, '')
            return str(self._cadfoldername) + str(path)
        else:
            return None
