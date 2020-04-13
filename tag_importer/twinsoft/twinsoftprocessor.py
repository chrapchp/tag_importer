import pandas as pd
import logging
import xml.etree.ElementTree as et
import re
import numpy as np

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
    TE_GROUP_NOT_FOUND = -106
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

    def __init__(self, message, errors):
        super().__init__(message)
        self.extended_error = errors


class TwinsoftProcessor:
    TW_MAX_TAG_LEN = 15
    TW_IGNORE_DATA = -9999
    TW_TAG_MAX_DESC_LEN = 50

    def __init__(self, xl_processor, twinsoft_tag_export_file, write_xml_file):
        self.xl_processor = xl_processor
        self.__logger = logging.getLogger(__name__)
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
            cnt = 1
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
                # return pd.DataFrame(tag_records).sort_values(by=['Group', 'ModbusAddress']).astype({'Signed': 'bool'},copy=True)
                x = pd.DataFrame(tag_records).sort_values(
                    by=['Group', 'ModbusAddress'])
                # print( 'here" {}'.format( x[x['Signed'].notnull()==False][['Signed','Tag']]))
                # x[x['Signed'].notnull()==False]['Signed2']=False
                x["Signed"].fillna("False", inplace=True)
                # x.Signed = x.Signed == 'True'
                # x['Signed'] = 'Peter'
                # x.to_clipboard()
                # force Empty Signed column to False
                # x = x.astype({'Signed': 'bool'},copy=True)
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

    def load_twinsoft_xml(self):
        self.__logger.info("Loading Twinsoft exported tag file...")
        self.__twinsoft_tags_df = self.__twinsoft_export_to_df(
            root_key="Tag", root_attrib_key='Name')

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

    def load_and_validate_memory_map(self):
        self.__logger.info("Loading Memory Map...")

        self.__xl_memory_map_df = self.xl_processor.memory_map_df
        df = self.__xl_memory_map_df.copy()
        # compute end address based on format type
        df['END_ADDRESS'] = np.where(df['FORMAT'].isin(['FLOAT', 'INT32', 'UINT32']),
                                     df['START_ADDRESS'] + df['LENGTH'] * 2 - 2, df['START_ADDRESS'] + df['LENGTH'] - 1)

        df['IS_BOOL'] = df['FORMAT'] == 'BOOL'
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
                                err_df = err_df.append({'ORIGIN_GROUP': row['GROUP'], 'ORIGIN_FORMAT': row['FORMAT'], 'ORIGIN_START_ADDRESS': row['START_ADDRESS'],
                                                        'ORIGIN_END_ADDRESS': row['END_ADDRESS'],
                                                        'CONFLICT_GROUP': nrow['GROUP'], 'CONFLICT_FORMAT': nrow['FORMAT'], 'CONFLICT_START_ADDR': int(nrow[
                                                            'START_ADDRESS']), 'CONFLICT_END_ADDR': int(nrow['END_ADDRESS'])
                                                        }, ignore_index=True)
        if err_df.shape[0] > 0:
            raise TwinsoftError("Memory Map Conflict:\n{}\n".format(
                err_df), TwinsoftError.TE_MEMORY_MAP_CONFLICT)

    def load_data(self):

        self.load_validate_tags()
        self.load_and_validate_memory_map()
        self.__logger.info("Loading Template...")
        self.__xl_template_df = self.xl_processor.template_df
        self.load_twinsoft_xml()

    def get_twinsoft_export_summary(self):
        """

        """
        if self.__twinsoft_tags_df['Group'].isnull().any():
            raise TwinsoftError(
                "One or more tags in Twinsoft are not part of group. Tags Must Belong to a GROUP for processing", TwinsoftError.TE_GROUP_EMPTY)

        x = self.__twinsoft_tags_df.groupby(['Group', 'Format', 'Signed']).agg({
            'ModbusAddress': ['min', 'max']})
        x.columns = ['MB_MIN', 'MB_MAX']

        x = x.reset_index()

        # one of the ways to get the merge to work with a TRUE/FALSE. Otherwise merge does not return correct value not both excel and xml signed are bool types
        # rather than object bool
        # print(self.__twinsoft_tags_df)

        x['Signed'] = x['Signed'] == 'True'
        x = x.astype({'Signed': 'bool', 'MB_MAX': int}, copy=True)
        self.__logger.debug("{}()\n{}".format(
            self.get_twinsoft_export_summary.__name__, x))
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
        if row['TS_FORMAT'] != 'DIGITAL':
            ret += '<Minimum>' + '0' + '</Minimum>\n'
            ret += '<Maximum>' + '1000' + '</Maximum>\n'
            ret += '<Resolution>' + '' + '</Resolution>\n'
        else:
            ret += '<Minimum />\n'
            ret += '<Maximum />\n'
            ret += '<Resolution />\n'
        ret += '<Group>' + row['FOLDER'] + '</Group>\n'
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

    def __validate_gen_df_addresses(self, df):
        merged_df = pd.merge(df, self.__xl_memory_map_df, left_on=[
            'Group', 'TS_FORMAT'], right_on=['GROUP', 'TS_FORMAT'], how='right')
        merged_df = merged_df[merged_df['Group'].notna()]

        merged_df['MAX_ADDRESS'] = np.where(merged_df['FORMAT_x'].isin(['FLOAT', 'INT32', 'UINT32']),
                                            merged_df['START_ADDRESS_y'] + merged_df['LENGTH_y'] * 2 - 2, merged_df['START_ADDRESS_y'] + merged_df['LENGTH_y'] - 1)
        merged_df.drop(['Group', 'Format', 'Signed', 'INITIAL_VALUE',  'GROUP_NUM_y',  'COMMENT_y',
                        'RULE', 'TS_SIGNED_x', 'LENGTH_x', 'MB_MIN', 'MB_MAX', 'CALC_INC', 'FOLDER', 'TS_FORMAT', 'LENGTH_y', 'START_ADDRESS_x', 'HAS_DATA', 'DESCRIPTION', 'GROUP_NUM_x', 'COMMENT_x', 'TS_SIGNED_y', 'FORMAT_y', 'SCRIPT_VALUE'], axis=1, inplace=True)

        merged_df.rename({'START_ADDRESS_y': 'MIN_ADDRESS',
                          'FORMAT_x': 'FORMAT'}, axis=1, inplace=True)
        errs = merged_df[(merged_df['CALC_ADDRESS'] > merged_df['MAX_ADDRESS']) | (
            merged_df['CALC_ADDRESS'] < merged_df['MIN_ADDRESS'])]
        if errs.shape[0] > 0:
            raise TwinsoftError("The following generated tags contain addresses that fall outside of the memory map.\n{}\n Revise MEMORY_MAP in excel file for flagged Group/Format.\n".format(
                errs), TwinsoftError.TE_CALC_ADDRESS_NOT_IN_MEMORY_MAP)

    def __validate_gen_df(self, gen_df):

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

        self.__validate_gen_df_addresses(gen_df)
        # check for duplicate calculated modbus addresses for DIGITALS and NON-DIGITAL TAGS

        subset_df = gen_df[gen_df['FORMAT'] == 'BOOL'].copy()

        dups = subset_df.duplicated(subset=['CALC_ADDRESS'])
        if dups.any():
            raise TwinsoftError('Duplicate BOOL addresses generated for the following: \n' + str(subset_df.loc[dups][[
                                'TAG', 'CALC_ADDRESS', 'FORMAT', 'FOLDER']]) + '\n', TwinsoftError.TE_DUPLICATE_BOOL_ADDR)

        subset_df = gen_df[gen_df['FORMAT'] != 'BOOL'].copy()
        dups = subset_df.duplicated(subset=['CALC_ADDRESS'])
        if dups.any():
            raise TwinsoftError('Duplicate ANALOG addresses generated for the following: \n' + str(subset_df.loc[dups][[
                                'TAG', 'CALC_ADDRESS', 'FORMAT', 'FOLDER']]) + '\n', TwinsoftError.TE_DUPLICATE_ANALOG_ADDR)

    def __generate_addressing(self, pending_tags_df):
        export_summary = self.get_twinsoft_export_summary()
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
                                   'Group', 'Format', 'Signed'], right_on=['FOLDER', 'TS_FORMAT', 'TS_SIGNED'], how='right')

        # gen_tags_merged.drop(
        #    ['Signed', 'Group', 'Format'], axis=1, inplace=True)
        # print(gen_tags_merged[['Group', 'Format', 'Signed',  'MB_MIN', 'MB_MAX', 'TAG', 'TS_FORMAT',
        #                       'TS_SIGNED', 'START_ADDRESS', 'LENGTH', 'FOLDER', 'FORMAT']])

        # print(gen_tags_merged.columns)
        gen_tags_merged['HAS_DATA'] = gen_tags_merged['Group'].isna() == False

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
            ['HAS_DATA', 'FOLDER', 'FORMAT']).cumcount()

        gen_tags_merged.loc[(gen_tags_merged['FORMAT'].isin(
            ['FLOAT', 'INT32', 'UINT32'])), 'CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'] * 2
        # gen_tags_merged['deleteme'] = gen_tags_merged['CALC_ADDRESS']
        gen_tags_merged['CALC_INC'] = 1
        gen_tags_merged.loc[(gen_tags_merged['FORMAT'].isin(
            ['FLOAT', 'INT32', 'UINT32'])), 'CALC_INC'] = 2
        gen_tags_merged['CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'] + \
            np.where(gen_tags_merged['HAS_DATA'] == True,
                     gen_tags_merged['MB_MAX'] + gen_tags_merged['CALC_INC'], gen_tags_merged['START_ADDRESS'])
        # np.where(gen_tags_merged['TS_FORMAT'] in ['FLOAT', 'INT32', 'UINT32'], 2, 1)

        gen_tags_merged['CALC_ADDRESS'] = gen_tags_merged['CALC_ADDRESS'].astype(
            int)

        self.__logger.debug("{}() - gen_tag_merged-dataframe\n{}".format(self.get_twinsoft_export_summary.__name__, gen_tags_merged[['Group', 'Format', 'Signed',  'MB_MIN', 'MB_MAX', 'TAG', 'TS_FORMAT',
                                                                                                                                     'TS_SIGNED', 'START_ADDRESS', 'LENGTH', 'FOLDER', 'FORMAT', 'CALC_ADDRESS', 'CALC_INC', 'HAS_DATA']]))
        self.__logger.debug("{}() - gen_tags_merged-data types\n{}".format(
            self.get_twinsoft_export_summary.__name__, gen_tags_merged.dtypes))

        self.__validate_gen_df(gen_tags_merged)
        self.__to_twinsoft_xml(gen_tags_merged)

    def generate_remote_tags(self, pattern):
        self.__logger.info("Remote Tag functionality not yet implemented.")

    def generate_tags(self, pattern):
        self.load_data()

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
            'GROUP', 'TYPE_y'], right_on=['GROUP', 'FORMAT'], how='left')

        errs = pattern_df[pattern_df['FORMAT_y'].isna()][[
            'GROUP', 'NEW_TAG', 'TYPE_y']]

        if errs.shape[0] > 0:
            raise TwinsoftError("GROUP not found in memory map. \n{} \nPossibly a TYPE in the TEMPLATE does not exist for a GROUP in the MEMORY_MAP.\n".format(
                errs), TwinsoftError.TE_GROUP_NOT_FOUND)

        # clean up headers
        pattern_df.drop(['CLASS', 'TAG_NAME', 'TAG_PATTERN', 'DESCRIPTION_x', 'TEMPLATE', 'GROUP', 'SUFFIX', 'TYPE_y',
                         'TYPE_x', 'FORMAT_x', 'DEVICE', 'IO', 'ADDRESS', 'DESCRIPTION_y', 'INITIAL_VALUE_x'], axis=1, inplace=True)
        pattern_df.rename({'NEW_TAG': 'TAG', 'NEW_DESC': 'DESCRIPTION', 'FORMAT_y': 'FORMAT',
                           'INITIAL_VALUE_y': 'INITIAL_VALUE'}, axis=1, inplace=True)

        exported_merged = pd.merge(self.__twinsoft_tags_df, pattern_df, left_on=[
            'Tag'], right_on=['TAG'])

        # self.__logger.debug("{}() - exported_merged-dataframe\n{}".format(self.generate_tags.__name__,exported_merged[['Tag','TAG']]))
        errs = list(exported_merged['Tag'])
        if len(errs):
            raise TwinsoftError('Generated tags ' + str(errs) + ' for pattern ' + pattern +
                                ' already exist in Twinsoft export file' + self.__twinsoft_tag_export_file, TwinsoftError.TE_TAGS_EXIST)

        self.__generate_addressing(pattern_df)

    def clone(self, tag_filter, group_filter, dest, address_offset, loop_no, replace_pattern):
        self.load_data()
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
            clone_df['Group'] = dest
        else:
            clone_df['Group'] = clone_df['Group'].str.replace(
                pat=replace_pattern, repl=loop_no,  regex=True)
        clone_df['Comment'] = clone_df['Comment'].str.replace(
            pat=replace_pattern, repl=loop_no, n=1, regex=True)

        clone_df.rename({'Tag': 'TAG', 'Comment': 'DESCRIPTION', 'Format': 'TS_FORMAT',
                         'Signed': 'TS_SIGNED', 'InitalValue': 'INITIAL_VALUE', 'Group': 'FOLDER', 'ModbusAddress': 'CALC_ADDRESS'}, axis=1, inplace=True)
        # print(clone_df[['TAG','DESCRIPTION','CALC_ADDRESS','FOLDER']])
        self.__to_twinsoft_xml(clone_df)
