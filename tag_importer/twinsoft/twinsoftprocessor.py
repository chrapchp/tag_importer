import pandas as pd
import logging
import xml.etree.ElementTree as et
import re
import sys
import numpy as np
from datetime import datetime
import verboselogs

from helpers import ExcelProcessor
from helpers import ExcelProcessorError

"""
Twinsoft tag export XML format

<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<TWinSoftTags>
  <Tag Name="DI1234567890123">
    <NewName>DI1234567890123</NewName>
    <Address>DIV00000</Address>
    <Format>DIGITAL</Format>
    <ModbusAddress>20480</ModbusAddress>
    <Comment>DICOMMENT12345678901234567890123456789012345678901</Comment>
    <InitalValue />
    <Signed />
    <TextTagSize />
    <Minimum />
    <Maximum />
    <Resolution />
    <Group>FOLDERNAME12345</Group>
    <Presentation Description="" StateOn="" StateOff="" Units="" NbrDecimals="">False</Presentation>
    <WriteAllowed WriteAllowed_Minimum="" WriteAllowed_Maximum="">False</WriteAllowed>
    <DisplayFormat>DECIMAL</DisplayFormat>
  </Tag>

"""

"""
Twinsoft remote tag export XML format

<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<TWinSoftRemoteTags>
  <RemoteTag TagName="RM2_1_DI0">
    <ExternalSource Device="RM2" Type="Coil" Address="20000" />
    <Operation Read="True" Quantity="16" />
    <Device DeviceType="RM2" ComPort="COM3" DeviceAddress="1" IPAddress="192.168.1.122" TriggerName="COM_TRIGGER" TriggerType="PositiveEdge" />
  </RemoteTag>

"""


class TwinsoftError(Exception):
    TE_PATTERN_NOT_FOUND = -100
    TE_TOO_MANY_PATTERNS = -101
    TE_XML_NOT_FOUND = -102
    TE_XML_ROOT_KEY_NOT_FOUND = -103
    TE_XML_ATTRIBUTE_KEY_NOT_FOUND = -104
    TE_TEMPLATE_NOT_FOUND = -105
    TE_MEM_ID_NOT_FOUND = -106
    TE_TAG_IN_EXPORT_FILE_EXIST = -107
    TE_TAGS_EXIST = -108
    TE_TAG_NAME_TOO_LONG = -109
    TE_TAG_DESC_TOO_LONG = -110
    TE_DUPLICATE_BOOL_ADDR = -111
    TE_DUPLICATE_ANALOG_ADDR = -112
    TE_DUPLICATE_TAG_NAME = -113
    TE_MEMORY_MAP_CONFLICT = -114
    TE_GROUP_EMPTY = -115
    TE_DOUBLE_UNDERSCORES = -116
    TE_CALC_ADDRESS_NOT_IN_MEMORY_MAP = -117
    TE_MAP_ENTRY_MISSING = -118
    TE_MAP_ENTRY_MEMORY_MAP_ENTRY_NOT_FOUND = -119
    TE_MAP_GROUP_TOO_LONG = -120

    def __init__(self, message, errors):
        super().__init__(message)
        self.extended_error = errors


class TwinsoftProcessor:
    TW_MAX_TAG_LEN = 15
    TW_MAX_GROUP_NAME_LEN = 15
    TW_IGNORE_DATA = -9999
    TW_TAG_MAX_DESC_LEN = 50

    def __init__(self, xl_processor, twinsoft_tag_export_file, write_xml_file):
        self.xl_processor = xl_processor
        self.__logger = verboselogs.VerboseLogger(__name__)
        self.__twinsoft_tag_export_file = twinsoft_tag_export_file
        self.__twinsoft_tags_df = None
        self.__xl_memory_map_df = None
        self.__xl_tags_df = None
        self.__xl_template_df = None
        self.__write_xml_file = write_xml_file
        self.__to_export_df = None

    def __twinsoft_export_to_df(self, root_key, root_attrib_key):
        """
        Convert twinsoft tag export in xml to pandas data frame

        The rootkey becomes a column name and root_attribute_key becomes the column name entry
        the rest of the XML use the element tag as the column name and element text as the value
        e.g.
        root_key = Tag, root_attrib_key = name
        <Taglist>
            <Tag name ='LS_001'>
                <Description>Level Switch</Description
                <ModbussAddress>1000</ModbusAddress>
            </Tag>
            <Tag name ='FI_001'>
                <Description>Flow Indicator</Description
                <ModbussAddress>1100</ModbusAddress>
            </Tag>
        </Taglist>
        becomes
        Tag     Description     ModbusAddress
        LS_001  Level Switch    1000
        FI_001  Flow Indicator  1100


        Parameters:
        root_key: str
                key to filter the xml document by. e.g. Tag
        root_attrib_key:  str
                root key name used extract the value of the key tag. e.g. Name

        Returns
        dataframe
                xml representation as a panda dataframe if lenth is > 0. TwinsoftError raised otherwise
        """
        try:
            xtree = et.parse(self.__twinsoft_tag_export_file)
            # self.__iter_tags( xtree )

            tag_records = []
            root = xtree.getroot()

            for tag_entry in root.findall(root_key):
                tag_record = {}
                tag_record[tag_entry.tag] = tag_entry.attrib[root_attrib_key]
                for tag_details in tag_entry:
                    # if (tag_details.tag == 'Signed') & (tag_details.text is None):
                    #    tag_record[tag_details.tag] = 'False'
                    #    print('*** {} value ={} for tag {}'.format(tag_details.tag,tag_record[tag_details.tag],tag_record[tag_entry.tag]))
                    # else:
                    tag_record[tag_details.tag] = tag_details.text
                    # print(tag_details.tag, tag_details.text, tag_details.attrib)
                # print("TAG REGCORD: {} {} ".format(cnt, tag_record))
                # cnt += 1
                tag_records.append(tag_record)

            if len(tag_records) == 0:
                raise TwinsoftError("<" + root_key + "> not found in " +
                                    self.__twinsoft_tag_export_file,  TwinsoftError.TE_XML_ROOT_KEY_NOT_FOUND)
            else:

                x = pd.DataFrame(tag_records).sort_values(
                    by=['Group', 'ModbusAddress'])

                x["Signed"].fillna("False", inplace=True)
                return x
        except FileNotFoundError:
            raise TwinsoftError("No such Twinsoft export file or directory " +
                                self.__twinsoft_tag_export_file, TwinsoftError.TE_XML_NOT_FOUND)
        except KeyError:
            raise TwinsoftError("Attribute " + root_attrib_key + " not found. e.g. <" + root_key + " " + root_attrib_key + "=> in file " +
                                self.__twinsoft_tag_export_file, TwinsoftError.TE_XML_ATTRIBUTE_KEY_NOT_FOUND)

    def replace_pattern(self, pattern, source, new_content):

        if re.search(pattern, source):
            pos = re.search(pattern, source).start()
            return source[:pos] + new_content + source[pos+1:]
        else:
            return source

    def load_twinsoft_xml(self, validate=True):
        self.__logger.info("Loading Twinsoft exported tag file...")

        self.__twinsoft_tags_df = self.__twinsoft_export_to_df(
            root_key="Tag", root_attrib_key='Name')

        #self.__twinsoft_tags_df = self.__twinsoft_tags_df.drop(self.__twinsoft_tags_df[self.__twinsoft_tags_df.TS_GROUP.isna()])
        self.__twinsoft_tags_df = self.__twinsoft_tags_df.astype(
            {'ModbusAddress': int}, copy=True)
        # A twinsoft group can belong to the same memory map entry. many to one
        map_df = self.xl_processor.tags_df[self.xl_processor.tags_df['CLASS'] == 'MAP']

        self.__twinsoft_tags_df = pd.merge(self.__twinsoft_tags_df, map_df, left_on=[
            'Group'], right_on=['TS_GROUP'], how='left')
        self.__twinsoft_tags_df.drop(
            ['CLASS', 'TAG_NAME', 'TAG_PATTERN', 'DESCRIPTION', 'TEMPLATE'],  axis=1, inplace=True) 

        if validate == True:
            t = self.__twinsoft_tags_df[self.__twinsoft_tags_df['TS_GROUP'].isna(
            )]
        
            if t.shape[0] > 0:
                raise TwinsoftError('\nMissing MAP Entry for \n {}\n Review TAGS sheet and ensure that entry exists and is correct CLASS=MAP is correct.\n'.format(
                    t.drop_duplicates(['Group'])[['Group']]), TwinsoftError.TE_MAP_ENTRY_MISSING)

            merged_df = pd.merge(self.__twinsoft_tags_df, self.xl_processor.memory_map_df, on=[
                'MEM_ID'], how='left')
            self.__logger.verbose("Memory Map Merge Columns: {}()\n{}".format(
                self.load_twinsoft_xml.__name__, merged_df.columns))
            self.__logger.verbose("Memory Map Merge Data: {}()\n{}".format(
                self.load_twinsoft_xml.__name__, merged_df[['Tag', 'TS_GROUP', 'TS_SIGNED', 'ModbusAddress', 'MEM_ID']]))
            t = merged_df[merged_df['MEM_TYPE'].isna()]
            if t.shape[0] > 0:
                raise TwinsoftError('\nMAP entry has no corresponding MEMORY_MAP Entry. \n {}\n Review TAGS sheet and ensure that entry exists and is correct CLASS=MAP is correct and/or MEMORY_MAP entries match.\n'.format(
                    t.drop_duplicates(['MEM_ID'])[['TS_GROUP', 'MEM_ID']]), TwinsoftError.TE_MAP_ENTRY_MISSING)

    def load_validate_tags(self):
        self.__logger.info("Loading Tags...")
        self.__xl_tags_df = self.xl_processor.tags_df

        # check if a duplicate tag name exists in TAGS sheet for Excel file with class BASE
        subset_df = self.__xl_tags_df[self.__xl_tags_df['CLASS'] == 'BASE'].copy(
        )
        subset_df = subset_df[subset_df['TAG_NAME'].notna()]
        dups = subset_df.duplicated(subset=['TAG_NAME'])
        if dups.any():
            raise TwinsoftError('Duplicate TAG_NAME defined in Sheet ' + ExcelProcessor.EXCEL_TAG_SHEET + ' in file ' + self.xl_processor.xl_file_name + '\n' + str(subset_df.loc[dups][[
                                'CLASS', 'TAG_NAME']]) + '\n', TwinsoftError.TE_DUPLICATE_TAG_NAME)

        # check if a duplicate tag name exists in TAG_PATTERN sheet for Excel file with class GENERATE
        subset_df = self.__xl_tags_df[self.__xl_tags_df['CLASS'] == 'GENERATE'].copy(
        )
        subset_df = subset_df[subset_df['TAG_PATTERN'].notna()]
        dups = subset_df.duplicated(subset=['TAG_PATTERN'])

        if dups.any():
            raise TwinsoftError('Duplicate TAG_NATAG_PATTERNME defined in Sheet ' + ExcelProcessor.EXCEL_TAG_SHEET + ' in file ' + self.xl_processor.xl_file_name + '\n' + str(subset_df.loc[dups][[
                                'CLASS', 'TAG_PATTERN']]) + '\n', TwinsoftError.TE_DUPLICATE_TAG_NAME)

    def load_and_validate_memory_map(self,ignore_errors):
        self.__logger.info("Loading Memory Map...")

        self.__xl_memory_map_df = self.xl_processor.memory_map_df
        df = self.__xl_memory_map_df.copy()
        # compute end address based on format type
        df['END_ADDRESS'] = np.where(df['MEM_TYPE'].isin(['FLOAT', 'INT32', 'UINT32']),
                                     df['START_ADDRESS'] + df['LENGTH'] * 2 - 2, df['START_ADDRESS'] + df['LENGTH'] - 1)

        df['IS_BOOL'] = df['MEM_TYPE'] == 'BOOL'
        grp = df.groupby(['IS_BOOL'])

        # df['overlap'] = (grp.apply(lambda x: ((x['START_ADDRESS'] <= x['END_ADDRESS'].shift(periods=-1, fill_value=0))
        #                                      & (x['START_ADDRESS'].shift(periods=-1, fill_value=0) <= x['END_ADDRESS'])))).reset_index(level=0, drop=True)

        # could not get a vectorized way to generate conflict map so ended up with iterating rows
        # look for overlaps in start/end addresses for BOOL and Non-BOOL groups

        err_df = pd.DataFrame(columns=['ORIGIN_GROUP', 'ORIGIN_FORMAT', 'ORIGIN_START_ADDRESS', 'ORIGIN_END_ADDRESS',
                                       'CONFLICT_GROUP', 'CONFLICT_FORMAT', 'CONFLICT_START_ADDR', 'CONFLICT_END_ADDR', ])
        for group_name, df_group in grp:
            row_iterator = df_group.iterrows()
            for i, row in row_iterator:
                if i != df.shape[0]-1:
                    df2 = df.shift(-1*(i+1)).copy()
                    df2.dropna(inplace=True)
                    for i, nrow in df2.iterrows():
                        if nrow['IS_BOOL'] == group_name:
                            # print("row['START_ADDRESS'] >= nrow['END_ADDRESS']:{}{} nrow['START_ADDRESS'] <= row['END_ADDRESS']{}{}".format(row['START_ADDRESS'],nrow['END_ADDRESS'],nrow['START_ADDRESS'],row['END_ADDRESS']))
                            # print('arg1 {}<={} arg3 {}<={}'.format(row['START_ADDRESS'],nrow['END_ADDRESS'],nrow['START_ADDRESS'],row['END_ADDRESS'] ))

                            if (row['START_ADDRESS'] <= nrow['END_ADDRESS']) & (nrow['START_ADDRESS'] <= row['END_ADDRESS']):
                                err_df = err_df.append({'ORIGIN_GROUP': row['MEM_ID'], 'ORIGIN_FORMAT': row['MEM_TYPE'], 'ORIGIN_START_ADDRESS': row['START_ADDRESS'],
                                                        'ORIGIN_END_ADDRESS': row['END_ADDRESS'],
                                                        'CONFLICT_GROUP': nrow['MEM_ID'], 'CONFLICT_FORMAT': nrow['MEM_TYPE'], 'CONFLICT_START_ADDR': int(nrow[
                                                            'START_ADDRESS']), 'CONFLICT_END_ADDR': int(nrow['END_ADDRESS'])
                                                        }, ignore_index=True)
        if not ignore_errors:                                                          
            if err_df.shape[0] > 0:
                raise TwinsoftError("Memory Map Conflict:\n{}\n".format(
                    err_df), TwinsoftError.TE_MEMORY_MAP_CONFLICT)

    def load_data(self, ignore_errors=False):

        self.load_validate_tags()
        self.load_and_validate_memory_map(ignore_errors)
        self.__logger.info("Loading Template...")
        self.__xl_template_df = self.xl_processor.template_df
        self.load_twinsoft_xml()

    def __as_memory_map(self, df):
 
        df.rename(columns={"MB_MIN": "START_ADDRESS",
                           "Format": "TS_FORMAT", "Signed": "TS_SIGNED"}, inplace=True)
        now = datetime.now()

        df['COMMENT'] = "Generated on: " + now.strftime("%d/%m/%Y %H:%M:%S")
        df['GROUP_NUM'] = 0
        df['MEM_TYPE'] = 'TODO'
        df.loc[df['TS_FORMAT'] == 'FLOAT', "MEM_TYPE"] = 'FLOAT'
        df.loc[df['TS_FORMAT'] == 'DIGITAL', "MEM_TYPE"] = 'BOOL'
        df.loc[(df['TS_FORMAT'] == '16BITS') & (
            df['TS_SIGNED'] == False), "MEM_TYPE"] = 'UINT16'
        df.loc[(df['TS_FORMAT'] == '16BITS') & (
            df['TS_SIGNED'] == True), "MEM_TYPE"] = 'INT16'
        df.loc[(df['TS_FORMAT'] == '32BITS') & (
            df['TS_SIGNED'] == False), "MEM_TYPE"] = 'UINT32'
        df.loc[(df['TS_FORMAT'] == '32BITS') & (
            df['TS_SIGNED'] == True), "MEM_TYPE"] = 'INT32'
        df.loc[(df['TS_FORMAT'] == 'BYTE') & (
            df['TS_SIGNED'] == False), "MEM_TYPE"] = 'UINT8'
        df['LENGTH'] = 0
        df = df.astype({'START_ADDRESS': int, 'MB_MAX': int}, copy=True)
        df.loc[df['MEM_TYPE'] == 'FLOAT', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS']) / 2
        df.loc[df['MEM_TYPE'] == 'UINT32', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS']) / 2
        df.loc[df['MEM_TYPE'] == 'INT32', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS']) / 2
        df.loc[df['MEM_TYPE'] == 'UINT16', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS'])
        df.loc[df['MEM_TYPE'] == 'INT16', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS'])
        df.loc[df['MEM_TYPE'] == 'UINT8', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS'])
        df.loc[df['MEM_TYPE'] == 'BOOL', "LENGTH"] = (
            df['MB_MAX'] - df['START_ADDRESS'])
        df['LENGTH'] = df['LENGTH'] + 1
        df = df.astype({'LENGTH': int}, copy=True)
        df = df[['MEM_ID', 'MEM_TYPE', 'START_ADDRESS', 'LENGTH',
                 'TS_FORMAT', 'TS_SIGNED', 'GROUP_NUM', 'COMMENT']]

        return df

    def get_twinsoft_export_summary(self, to_memory_map=False, root_tags=False):
        """

        """
        if not root_tags:
            if self.__twinsoft_tags_df['Group'].isnull().any():
                raise TwinsoftError(
                    "One or more tags in Twinsoft are not part of group. Tags Must Belong to a GROUP for processing. Try --root_tags if root tags exist", TwinsoftError.TE_GROUP_EMPTY)

        x = self.__twinsoft_tags_df.groupby(['Group', 'Format', 'Signed']).agg({
            'ModbusAddress': ['min', 'max']})
        x.columns = ['MB_MIN', 'MB_MAX']

        x = x.reset_index()

        # one of the ways to get the merge to work with a TRUE/FALSE. Otherwise merge does not return correct value not both excel and xml signed are bool types
        # rather than object bool
        # print(self.__twinsoft_tags_df)

        x['Signed'] = x['Signed'] == 'True'
        x = x.astype({'Signed': 'bool', 'MB_MAX': int}, copy=True)
        # A twinsoft group can belong to the same memory map entry. many to one

        map_df = self.xl_processor.tags_df[self.xl_processor.tags_df['CLASS'] == 'MAP']
        #map_df = self.xl_processor.tags_df.drop_duplicates(['TS_GROUP'])

        x = pd.merge(x, map_df, left_on=[
            'Group'], right_on=['TS_GROUP'], how='left')
        x.drop(['CLASS', 'TAG_NAME',  'TAG_PATTERN',
                'DESCRIPTION', 'TEMPLATE','ADDRESS','TAG_INITIAL_VALUE','TAG_TYPE'],  axis=1, inplace=True)
        self.__logger.verbose("{}()\n{}".format(
            self.get_twinsoft_export_summary.__name__, x))
        if to_memory_map == True:
            x = self.__as_memory_map(x)
        return x

    def __xml_encode_tag(self, row):
        ret = '<Tag Name=\"' + row['TAG'] + '\">\n'
        ret += '<NewName>' + row['TAG'] + '</NewName>\n'
        ret += '<Address />\n'
        ret += '<Format>' + row['TS_FORMAT'] + '</Format>\n'
        ret += '<ModbusAddress>' + \
            str(row['CALC_ADDRESS']) + '</ModbusAddress>\n'
        ret += '<Comment>' + row['DESCRIPTION'] + '</Comment>\n'
        if row['INITIAL_VALUE'] != TwinsoftProcessor.TW_IGNORE_DATA:
            ret += '<InitalValue>' + \
                str(row['INITIAL_VALUE']) + '</InitalValue>\n'
        else:
            ret += '<InitalValue />\n'
        if row['TS_FORMAT'] != 'DIGITAL':
            ret += '<Signed>' + str(row['TS_SIGNED']) + '</Signed>\n'
        else:
            ret += '<Signed />\n'
        ret += '<TextTagSize />\n'
        if row['TS_FORMAT'] == 'TEXT':
            ret += '<TextTagSize>' + str(row['TEXT_LEN']) + '</TextTagSize>'
        elif row['TS_FORMAT'] != 'DIGITAL':
            ret += '<Minimum>' + '0' + '</Minimum>\n'
            ret += '<Maximum>' + '1000' + '</Maximum>\n'
            ret += '<Resolution>' + '' + '</Resolution>\n'
       
        else:
            ret += '<Minimum />\n'
            ret += '<Maximum />\n'
            ret += '<Resolution />\n'
        ret += '<Group>' + row['TS_GROUP'] + '</Group>\n'
        ret += '<Presentation Description=\"\" StateOn=\"\" StateOff=\"\" Units=\"\" NbrDecimals=\"\">False</Presentation>'
        ret += '<WriteAllowed WriteAllowed_Minimum=\"\" WriteAllowed_Maximum=\"\">False</WriteAllowed>'
        ret += '<DisplayFormat>DECIMAL</DisplayFormat>'
        ret += '</Tag>'
        return ret

    def __to_twinsoft_xml(self, gen_df):
        # to_export = pd.DataFrame().reindex_like(self.__twinsoft_tags_df)

        with open(self.__write_xml_file, 'w') as xmlFile:
            xmlFile.write('<TWinSoftTags>\n')
            xmlFile.write('\n'.join(gen_df.apply(
                self.__xml_encode_tag, axis=1)))
            xmlFile.write('</TWinSoftTags>')

    def __validate_gen_df_addresses(self, df, blind_validation=False):

        merged_df = pd.merge(df, self.__xl_memory_map_df, left_on=[
            'MEM_ID', 'TS_FORMAT', 'TS_SIGNED'], right_on=['MEM_ID', 'TS_FORMAT', 'TS_SIGNED'], how='left')

        merged_df = merged_df[merged_df['MEM_TYPE_x'].notna()]
        if not blind_validation:
            if merged_df.shape[0] == 0:
                raise TwinsoftError(
                    "One or more entries for MEM_ID {} not found in memory map during validation of generated modbus addresses. \n "
                    "If Cloning, ensure mapping exists in MEMORY_MAP. Passing the --blind_validation as command option forces no address validation "
                    "and this check will be bypassed.\n".format(df['MEM_ID'].unique()), TwinsoftError.TE_MEM_ID_NOT_FOUND)

        merged_df['MAX_ADDRESS'] = np.where(merged_df['MEM_TYPE_x'].isin(['FLOAT', 'INT32', 'UINT32']),
                                            merged_df['START_ADDRESS_y'] + merged_df['LENGTH_y'] * 2 - 2, merged_df['START_ADDRESS_y'] + merged_df['LENGTH_y'] - 1)
        # merged_df.drop(['Group', 'Format', 'Signed', 'INITIAL_VALUE',  'GROUP_NUM_y',  'COMMENT_y',
        #                 'TS_SIGNED_x', 'LENGTH_x', 'MB_MIN', 'MB_MAX', 'CALC_INC',  'TS_FORMAT', 'LENGTH_y', 'START_ADDRESS_x', 'HAS_DATA', 'DESCRIPTION', 'GROUP_NUM_x', 'COMMENT_x', 'TS_SIGNED_y', 'MEM_TYPE_y', 'SCRIPT_VALUE'], axis=1, inplace=True)

        merged_df.rename({'START_ADDRESS_y': 'MIN_ADDRESS',
                          'FORMAT_x': 'FORMAT'}, axis=1, inplace=True)
        errs = merged_df[(merged_df['CALC_ADDRESS'] > merged_df['MAX_ADDRESS']) | (
            merged_df['CALC_ADDRESS'] < merged_df['MIN_ADDRESS'])]

        if errs.shape[0] > 0:
            raise TwinsoftError("The following generated tags contain addresses that fall outside of the memory map.\n{}\n Revise MEMORY_MAP in excel file for flagged Group/Format.\n".format(
                errs[['TAG',  'MEM_ID', 'MEM_TYPE_x', 'MIN_ADDRESS', 'CALC_ADDRESS', 'MAX_ADDRESS']]), TwinsoftError.TE_CALC_ADDRESS_NOT_IN_MEMORY_MAP)

    def __group_too_long_count(self, group_entry):
        return len([group_element for group_element in group_entry.split("\\", -1) if len(group_element) > TwinsoftProcessor.TW_MAX_GROUP_NAME_LEN])

    def __validate_gen_df(self, gen_df, blind_validation=False):

        t = gen_df.loc[(gen_df['TAG'].str.contains('.+__.+', regex=True))]
        if t.shape[0] > 0:
            raise TwinsoftError('\nGenerated Tag Names: \n' + str(list(
                t['TAG'])) + '\n contains two consecutive underscores (__). Twinsoft will not import them. \nMost likely a GENERATE entry with a TAG_PATTERN defined with _* and the template SUFFIX has _X ', TwinsoftError.TE_DOUBLE_UNDERSCORES)
        # check if any of the generated tag names > MAX permitted
        t = gen_df.loc[(gen_df['TAG'].str.len() >
                        TwinsoftProcessor.TW_MAX_TAG_LEN)]
        if t.shape[0] > 0:
            raise TwinsoftError('\nGenerated Tag Names: \n' + str(list(t['TAG'])) + '\ngreather than ' + str(
                TwinsoftProcessor.TW_MAX_TAG_LEN) + ' characters. Consider shortening the template suffix or tag prefix.', TwinsoftError.TE_TAG_NAME_TOO_LONG)
        # check if any of the generated tag descriptions > MAX permitted
        t = gen_df.loc[(gen_df['DESCRIPTION'].str.len() >
                        TwinsoftProcessor.TW_TAG_MAX_DESC_LEN)]
        if t.shape[0] > 0:
            raise TwinsoftError('\nGenerated Descriptions for Tag Names \n' + str(list(t['TAG'])) + '\ngreather than ' + str(
                TwinsoftProcessor.TW_TAG_MAX_DESC_LEN) + ' characters. Consider shortening the template description or tag description prefix.', TwinsoftError.TE_TAG_DESC_TOO_LONG)

        # check if group and sub groups lengths are too long

        gen_df['GROUP_ERR_CNT'] = gen_df.apply(
            lambda x: self.__group_too_long_count(x['TS_GROUP']), axis=1)

        t = gen_df.loc[gen_df['GROUP_ERR_CNT'] != 0]
        if t.shape[0] > 0:
            raise TwinsoftError('\n Group Names {0} too long \n Max is {1}\n'.format(t['TS_GROUP'].unique(
            ), TwinsoftProcessor.TW_MAX_GROUP_NAME_LEN), TwinsoftError.TE_MAP_GROUP_TOO_LONG)
        gen_df.drop(['GROUP_ERR_CNT'], axis=1, inplace=True)

        self.__validate_gen_df_addresses(gen_df, blind_validation)
        # check for duplicate calculated modbus addresses for DIGITALS and NON-DIGITAL TAGS

        subset_df = gen_df[gen_df['MEM_TYPE'] == 'BOOL'].copy()

        dups = subset_df.duplicated(subset=['CALC_ADDRESS'])
        if dups.any():
            raise TwinsoftError('Duplicate BOOL addresses generated for the following: \n' + str(subset_df.loc[dups][[
                                'TAG', 'CALC_ADDRESS', 'MEM_TYPE', 'TS_GROUP_x']]) + '\n', TwinsoftError.TE_DUPLICATE_BOOL_ADDR)

        subset_df = gen_df[gen_df['MEM_TYPE'] != 'BOOL'].copy()
        dups = subset_df.duplicated(subset=['CALC_ADDRESS'])
        if dups.any():
            #print("ere")
            #print(subset_df.loc[dups])
            raise TwinsoftError('Duplicate ANALOG addresses generated for the following: \n' + str(subset_df.loc[dups][[
                                'TAG', 'CALC_ADDRESS', 'MEM_TYPE', 'TS_GROUP']]) + '\n', TwinsoftError.TE_DUPLICATE_ANALOG_ADDR)

    def __generate_addressing(self, pending_tags_df):
        export_summary = self.get_twinsoft_export_summary()
        export_summary.dropna(axis=0, inplace=True)
        
        #
        # e.g. entry 0 requries a tag XY_110_OCA to be created with template starting at 1400 but CHAMBER 1\LOCALS has a tags starting at 1700
        #       the new tag will be generated starting at 1837 + 1 rather than 1400
        #               Group   Format Signed MB_MIN MB_MAX             TAG TS_FORMAT  TS_SIGNED  START_ADDRESS  LENGTH            FOLDER  FORMAT
        # 0   CHAMBER 1\LOCALS   16BITS  False   1700   1837      XY_110_OCA    16BITS      False           1400     100  CHAMBER 1\LOCALS  UINT16
        # 1   CHAMBER 1\LOCALS   16BITS  False   1700   1837     KX_126_OFAP    16BITS      False           1400     100  CHAMBER 1\LOCALS  UINT16
        # C11_PRIMARY_CC tag does not have a folder in the export called COMMUNICATION so it will start at 570
        # 24               NaN      NaN    NaN    NaN    NaN  C11_PRIMARY_CC    16BITS      False            570      20     COMMUNICATION  UINT16
        # Same for C11_PRIMARY_ST and it will start at 570 + 1
        # 25               NaN      NaN    NaN    NaN    NaN  C11_PRIMARY_ST    16BITS      False            570      20     COMMUNICATION  UINT16

        gen_tags_merged = pd.merge(export_summary, pending_tags_df, left_on=[
                                   'MEM_ID', 'Format', 'Signed'], right_on=['MEM_ID', 'TS_FORMAT', 'TS_SIGNED'], how='right')

        gen_tags_merged['HAS_DATA'] = gen_tags_merged['Group'].isna() == False

        max_merged = gen_tags_merged[gen_tags_merged['HAS_DATA'] == True].copy(
        )

        idx = max_merged.groupby(['MEM_ID', 'MEM_TYPE'])[
            'MB_MAX'].transform(max) == max_merged['MB_MAX']

        gen_tags_merged = gen_tags_merged[gen_tags_merged['HAS_DATA'] == False]
        max_merged = max_merged[idx]
        gen_tags_merged = pd.concat(
            [max_merged, gen_tags_merged], ignore_index=True)

        if (gen_tags_merged['HAS_DATA'] == True).shape[0] > 0:
            self.__logger.warning(" Twinsoft Tag Export File {} containts tag definitions. No checks for addressing conflict are made if tags don't exist in folder defined in the memory map. May run into Twinsoft import errors.".format(
                self.__twinsoft_tag_export_file))

        # count up tags by grouping existing folders in twinsoft then
        #               Group   Format Signed MB_MIN  MB_MAX             TAG TS_FORMAT  TS_SIGNED  START_ADDRESS  LENGTH            FOLDER  FORMAT  CALC_ADDRESS  CALC_INC  HAS_DATA
        # 0   CHAMBER 1\LOCALS   16BITS  False   1700  1837.0      XY_110_OCA    16BITS      False           1400     100  CHAMBER 1\LOCALS  UINT16          1838         1      True
        # 1   CHAMBER 1\LOCALS   16BITS  False   1700  1837.0     KX_126_OFAP    16BITS      False           1400     100  CHAMBER 1\LOCALS  UINT16          1839         1      True
        # 2   CHAMBER 1\LOCALS   16BITS  False   1700  1837.0      KX_126_OFH    16BITS      False           1400     100  CHAMBER 1\LOCALS  UINT16          1840         1      True
        #                                        ....
        # 45               NaN      NaN    NaN    NaN     NaN      ALW_DD_FLT    32BITS       True            350      50           GLOBALS   INT32           350         2     False
        # 46               NaN      NaN    NaN    NaN     NaN     ALW_DD_FLT4    32BITS      False            450      50           GLOBALS  UINT32           450         2     False
        # 47               NaN      NaN    NaN    NaN     NaN    ALW_DD_FLT23    32BITS      False            450      50           GLOBALS  UINT32           452         2     False

        gen_tags_merged['CALC_ADDRESS'] = gen_tags_merged.groupby(
            ['HAS_DATA', 'MEM_ID', 'MEM_TYPE']).cumcount()

        gen_tags_merged.loc[(gen_tags_merged['MEM_TYPE'].isin(
            ['FLOAT', 'INT32', 'UINT32'])), 'CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'] * 2
        gen_tags_merged.loc[(gen_tags_merged['MEM_TYPE'].isin(
            ['TEXT'])), 'CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'] *  gen_tags_merged['TEXT_LEN']     
          
        # gen_tags_merged['deleteme'] = gen_tags_merged['CALC_ADDRESS']
        gen_tags_merged['CALC_INC'] = 1
        gen_tags_merged.loc[(gen_tags_merged['MEM_TYPE'].isin(
            ['TEXT'])), 'CALC_INC'] =  gen_tags_merged['TEXT_LEN']  
        
        gen_tags_merged.loc[(gen_tags_merged['MEM_TYPE'].isin(
            ['FLOAT', 'INT32', 'UINT32'])), 'CALC_INC'] = 2

        gen_tags_merged['CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'] + \
            np.where(gen_tags_merged['HAS_DATA'] == True,
                     gen_tags_merged['MB_MAX'] + gen_tags_merged['CALC_INC'], gen_tags_merged['START_ADDRESS'])


        # np.where(gen_tags_merged['TS_FORMAT'] in ['FLOAT', 'INT32', 'UINT32'], 2, 1)

        gen_tags_merged['CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'].astype(
            int)

        self.__logger.verbose("{}() - gen_tag_merged-dataframe\n{}".format(self.get_twinsoft_export_summary.__name__, gen_tags_merged[['Group', 'Format', 'Signed',  'MB_MIN', 'MB_MAX', 'TAG', 'TS_FORMAT',
                                                                                                                                     'TS_SIGNED', 'START_ADDRESS', 'LENGTH', 'MEM_ID', 'MEM_TYPE','CALC_ADDRESS', 'CALC_INC', 'HAS_DATA']]))
        self.__logger.verbose("{}() - gen_tags_merged-data types\n{}".format(
            self.get_twinsoft_export_summary.__name__, gen_tags_merged.dtypes))

        gen_tags_merged.rename(
            {'TS_GROUP_y': 'TS_GROUP'}, axis=1, inplace=True)
        self.__validate_gen_df(gen_tags_merged)

        self.__to_twinsoft_xml(gen_tags_merged)

    def generate_remote_tags(self, pattern):
        self.__logger.info("Remote Tag functionality not yet.")

    def generate_tags(self, pattern, ignore_map_errors):
        self.load_data(ignore_map_errors)

        self.__logger.info("Generatings Tags for pattern <" + pattern + ">...")
        pattern_df = self.__xl_tags_df[(self.__xl_tags_df.CLASS == 'GENERATE') & (
            self.__xl_tags_df.TAG_PATTERN.str.contains(pattern, regex=True))]

        # join tag list and templates and generate which will contain tag_pattern and suffix
        pattern_df = pd.merge(
            pattern_df, self.__xl_template_df, on='TEMPLATE', how='left')

        # check if any templates are not found and abort process if any entry does not line up
        errs = list(pattern_df[pattern_df['SUFFIX'].isna()]
                    ['TEMPLATE'].unique())

        if len(errs):
            raise TwinsoftError('TEMPLATES ' + str(errs) + ' defined in sheet ' + ExcelProcessor.EXCEL_TAG_SHEET + ' not found under ' +
                                ExcelProcessor.EXCEL_TEMPLATE + ' sheet for file ' + self.xl_processor.xl_file_name, TwinsoftError.TE_TEMPLATE_NOT_FOUND)
        # do generate tags and descriptions
        try:

            pattern_df["NEW_TAG"] = pattern_df.apply(
                lambda x: self.replace_pattern(r'\*', x['TAG_PATTERN'], x['SUFFIX']), axis=1)
            pattern_df["NEW_DESC"] = pattern_df.apply(lambda x: self.replace_pattern(
                r'\*', x['DESCRIPTION_x'], x['DESCRIPTION_y']), axis=1)
        except ValueError:
            raise TwinsoftError("Pattern " + pattern + " does not exist in sheet " + ExcelProcessor.EXCEL_TAG_SHEET +
                                " in file " + self.xl_processor.xl_file_name, TwinsoftError.TE_PATTERN_NOT_FOUND)
        # for each entry we need to get the adress ranges by group and format
        pattern_df = pd.merge(pattern_df, self.__xl_memory_map_df, left_on=[
            'MEM_ID', 'TYPE'], right_on=['MEM_ID', 'MEM_TYPE'], how='left')

        errs = pattern_df[pattern_df['MEM_TYPE'].isna()][[
            'MEM_ID', 'NEW_TAG', 'TYPE']]

        if errs.shape[0] > 0:
            raise TwinsoftError("MEM_ID not found in memory map. \n{} \nPossibly a TYPE in the TEMPLATE does not exist for a GROUP in the MEMORY_MAP.\n".format(
                errs), TwinsoftError.TE_MEM_ID_NOT_FOUND)

        # clean up headers
        pattern_df.drop(['CLASS', 'TAG_NAME', 'TAG_PATTERN', 'DESCRIPTION_x', 'TEMPLATE',  'SUFFIX', 'TYPE',
                         'DESCRIPTION_y'], axis=1, inplace=True)
        pattern_df.rename({'NEW_TAG': 'TAG', 'NEW_DESC': 'DESCRIPTION', 'FORMAT_y': 'FORMAT',
                           'INITIAL_VALUE_y': 'INITIAL_VALUE'}, axis=1, inplace=True)

        exported_merged = pd.merge(self.__twinsoft_tags_df, pattern_df, left_on=[
            'Tag'], right_on=['TAG'])

        # self.__logger.verbose("{}() - exported_merged-dataframe\n{}".format(self.generate_tags.__name__,exported_merged[['Tag','TAG']]))
        errs = list(exported_merged['Tag'])
        if len(errs):
            raise TwinsoftError('Generated tags ' + str(errs) + ' for pattern ' + pattern +
                                ' already exist in Twinsoft export file' + self.__twinsoft_tag_export_file, TwinsoftError.TE_TAGS_EXIST)

        self.__generate_addressing(pattern_df)

    def clone(self, tag_filter, group_filter, dest, address_offset, loop_no, replace_pattern, blind_validation, group_pattern=None, group_replace=None,ignore_map_errors=None):
        self.load_data(ignore_map_errors)
        df = self.__twinsoft_tags_df

        clone_df = df[(df['Tag'].str.contains(tag_filter, regex=True)) & (
            df['Group'].str.contains(group_filter, regex=True))].copy()
        clone_df['Address'] = None

        clone_df['InitalValue'] = np.where(clone_df['InitalValue'].isnull(
        ), TwinsoftProcessor.TW_IGNORE_DATA, clone_df['InitalValue'])
        clone_df = clone_df.astype({'ModbusAddress': int}, copy=True)
        # print(clone_df.dtypes)
        clone_df['ModbusAddress'] += address_offset
        clone_df['Tag'] = clone_df['Tag'].str.replace(
            pat=replace_pattern, repl=loop_no, n=1, regex=True)
        clone_df['NewName'] = clone_df['NewName'].str.replace(
            pat=replace_pattern, repl=loop_no, n=1, regex=True)

        if dest is not None:
            clone_df['TS_GROUP'] = dest
        elif group_pattern is None:
            clone_df['TS_GROUP'] = clone_df['TS_GROUP'].str.replace(
                pat=replace_pattern, repl=loop_no,  regex=True)
        else:
            clone_df['TS_GROUP'] = clone_df['TS_GROUP'].str.replace(
                pat=group_pattern, repl=group_replace,  regex=True)

        clone_df['Comment'] = clone_df['Comment'].str.replace(
            pat=replace_pattern, repl=loop_no, n=1, regex=True)
        clone_df['MEM_ID'] = clone_df['MEM_ID'].str.replace(
            pat=replace_pattern, repl=loop_no, n=1, regex=True)

        #clone_df['FOLDER'] = clone_df['Group']
        clone_df.rename({'Tag': 'TAG', 'Comment': 'DESCRIPTION', 'Format': 'TS_FORMAT',
                         'Signed': 'TS_SIGNED', 'InitalValue': 'INITIAL_VALUE',  'ModbusAddress': 'CALC_ADDRESS'}, axis=1, inplace=True)

        clone_df['TS_SIGNED'] = clone_df['TS_SIGNED'] == 'True'
        clone_df = clone_df.astype({'TS_SIGNED': 'bool'}, copy=True)
        clone_df = pd.merge(clone_df, self.__xl_memory_map_df, left_on=[
            'MEM_ID', 'TS_FORMAT', 'TS_SIGNED'], right_on=['MEM_ID', 'TS_FORMAT', 'TS_SIGNED'], how='left')
        self.__logger.verbose("Clone df: {}()\n{}".format(self.clone.__name__, clone_df[[
                            'TAG', 'CALC_ADDRESS', 'TS_FORMAT', 'TS_SIGNED', 'MEM_ID']]))
        if clone_df.shape[0] == 0:
            raise TwinsoftError("tag_filter: {0} and/or group_filter: {1} did not find anything to clone.\n".format(
                tag_filter, group_filter), TwinsoftError.TE_PATTERN_NOT_FOUND)
        if not ignore_map_errors:
            self.__validate_gen_df(clone_df, blind_validation=blind_validation)
        self.__to_twinsoft_xml(clone_df)

    def create(self, tag_filter, group_filter):
        self.load_data()
        #df = self.__twinsoft_tags_df

        create_df = self.__xl_tags_df[(self.__xl_tags_df.CLASS == 'BASE') & (self.__xl_tags_df['TAG_NAME'].str.contains(
            tag_filter, regex=True)) & (self.__xl_tags_df['TS_GROUP'].str.contains(group_filter, regex=True))].copy()


        # join tag list and memory 
        create_df = pd.merge(create_df, self.__xl_memory_map_df, left_on=['MEM_ID','TAG_TYPE'], right_on=['MEM_ID','MEM_TYPE'], how='left')
        errs = create_df[create_df['MEM_TYPE'].isna()]
        if errs.shape[0] > 0:
            raise TwinsoftError('Mem_ID for tags \n{}\n not found or missing in MEMORY_MAP\n'.format(errs[['TAG_NAME','TS_GROUP','MEM_ID']]), TwinsoftError.TE_TAGS_EXIST)

        errs = pd.merge(self.__twinsoft_tags_df, create_df, left_on=['Tag'], right_on=['TAG_NAME'])

        if errs.shape[0]>0:
            raise TwinsoftError('\nTags: \n{} \nalready exist in Twinsoft export file\n'.format(errs['TAG_NAME'].to_string()), TwinsoftError.TE_TAGS_EXIST)
        
        create_df.rename({'TAG_NAME': 'TAG','TAG_INITIAL_VALUE': 'INITIAL_VALUE'}, axis=1, inplace=True)     
        #print(create_df)  
        self.__generate_addressing(create_df) 
    

        #print(create_df)
        #self.__validate_gen_df(create_df)
