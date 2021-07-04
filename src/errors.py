class Error(Exception):
    pass

class StartNotFound(Error):
    pass

class StopNotFound(Error):
    pass

class TableCropError(Error):
    pass

class ExcelDecodeError(Error):
    pass

class EmptyTableError(Error):
    pass

class FileTypeError(Error):
    pass