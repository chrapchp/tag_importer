import pandas as pd
from xlrd import XLRDError
import logging
import sys


class ExcelProcessorError(Exception):
    EE_TAB_NOT_FOUND = -200
    EE_FILE_NOT_FOUND = -201
    EE_TAB_EMPTY = -203
    EE_EMPTY_CELLS = -204

    def __init__(self, message, errors=0):
        super().__init__(message)
        self.extended_error = errors


class ExcelProcessor:
    EXCEL_TAG_SHEET = 'TAGS'
    EXCEL_MEMORY_MAP_SHEET = 'MEMORY_MAP'
    EXCEL_TEMPLATE = 'TEMPLATE'
    EXCEL_REMOTE_TAGS = 'REMOTE_TAGS'
    EXCEL_REMOTE_DEVICES = 'REMOTE_DEVICES'    

    def __init__(self, xl_file_name):

        self.xl_file_name = xl_file_name
        self.__logger = logging.getLogger(__name__)
        self.__tags_df = None
        self.__template_df = None
        self.__memory_map_df = None
        self.__remote_devices_df = None
        self.__remote_tags_df = None


    def __open_tab(self, tab_name, empty_tab_allowed=False, empty_cells_allowed=True):
        try:
            df = pd.read_excel(self.xl_file_name, sheet_name=tab_name)
            if (empty_tab_allowed == False) & (len(df) == 0):
                raise ExcelProcessorError("Tab name " + tab_name + " in " + self.xl_file_name +
                                          " cannot be empty.", ExcelProcessorError.EE_TAB_EMPTY)
            if (empty_cells_allowed == False) & (df.isnull().sum().sum() > 0):
                raise ExcelProcessorError("Tab name " + tab_name + " in " + self.xl_file_name +
                                          " has 1 more empty cells. All cells must have a value.", ExcelProcessorError.EE_EMPTY_CELLS)
            return df
        except FileNotFoundError:
            raise ExcelProcessorError(
                "No such Excel file or directory " + self.xl_file_name, ExcelProcessorError.EE_FILE_NOT_FOUND)
        except XLRDError:
            raise ExcelProcessorError(
                "Tab [" + tab_name + "] not found in file " + self.xl_file_name, ExcelProcessorError.EE_TAB_NOT_FOUND)

    @property
    def tags_df(self):
        if self.__tags_df is None:
            self.__tags_df = self.__open_tab(ExcelProcessor.EXCEL_TAG_SHEET)
        return self.__tags_df

    @property
    def template_df(self):
        if self.__template_df is None:
            self.__template_df = self.__open_tab(
                ExcelProcessor.EXCEL_TEMPLATE, empty_cells_allowed=False)
        return self.__template_df

    @property
    def memory_map_df(self):
        if self.__memory_map_df is None:
            self.__memory_map_df = self.__open_tab(
                ExcelProcessor.EXCEL_MEMORY_MAP_SHEET, empty_cells_allowed=False)

        return self.__memory_map_df

    @property
    def remote_devices_df(self):
        if self.__remote_devices_df is None:
            self.__remote_devices_df = self.__open_tab(
                ExcelProcessor.EXCEL_REMOTE_DEVICES, empty_cells_allowed=False)

        return self.__remote_devices_df


    @property
    def remote_tags_df(self):
        if self.__remote_tags_df is None:
            self.__remote_tags_df = self.__open_tab(
                ExcelProcessor.EXCEL_REMOTE_TAGS, empty_cells_allowed=False)

        return self.__remote_tags_df        

