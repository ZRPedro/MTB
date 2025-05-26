from enum import Enum


class ResultType(Enum):
    RMS       = 0 #PowerFactory standard output
    EMT_INF   = 1 #PSCAD legacy .inf/.csv support
    EMT_PSOUT = 2 #PSCAD .psout
    EMT_CSV   = 3 #PSCAD .psout -> .csv support
    EMT_ZIP   = 4 #PSCAD .psout -> .zip, .gz, .bz2 and .xz support


class Result:
    def __init__(self, typ : ResultType, rank : int, projectName : str, bulkname : str, fullpath : str, group : str) -> None:
        self.typ = typ
        self.rank = rank
        self.projectName = projectName
        self.bulkname = bulkname
        self.fullpath = fullpath
        self.group = group
        self.shorthand = f'{group}\\{projectName}'
