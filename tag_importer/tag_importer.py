
import pandas as pd
import argparse
import sys
import logging

import click


from twinsoft import TwinsoftProcessor
from twinsoft import TwinsoftError
from helpers import ExcelProcessor
from helpers import ExcelProcessorError


logger = logging.getLogger(__name__)


@click.group()
@click.option('--excel', required=True, help='Excel file containing tags and memory map')
@click.option('--xmlin', required=True, help='Exported tag XML file from Twinsoft')
@click.option('--xmlout', required=True, help='Output file of generated XML file')
@click.option('--verbose', is_flag=True, help="Will print more messages")
@click.pass_context
def main(ctx, excel, xmlin, xmlout, verbose):
    ctx.ensure_object(dict)
    ctx.obj['excel'] = excel
    ctx.obj['twinsoft_export'] = xmlin
    ctx.obj['xmlout'] = xmlout
    ctx.obj['verbose'] = verbose
    ctx.obj['excel_processor'] = ExcelProcessor(excel)

    ctx.obj['twinsoft_processor'] = TwinsoftProcessor(
        ctx.obj['excel_processor'], xmlin, xmlout)

    if verbose:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(
            format='%(asctime)s %(levelname)s- %(message)s', level=logging.INFO)
        # logging.basicConfig(level=logging.INFO)


@main.command()
@click.option('--pattern', required=True, help='Generate Tags for a given TAG_PATTERN defined in excel TAGS tab')
@click.pass_context
def generate(ctx, pattern):
    '''
    Generate tags using pattern defined in XL
    '''
    if pattern == '?':
        logger.info(
            'Patterns use RegEx syntax.  Samples are shown below. Patterns are shown between the single quotes: \'\'')
        logger.info(
            '\'.+\'                 -> selects all TAG_PATTERNS under GENERATE class')
        logger.info(
            '\'^[A-Z]\d{2}_.+\'     -> Start with 1 Capital letter followed by 2 digits and underscore, then 1 or more characters e.g. C12_PRIMARY*')
        logger.info(
            '\'^C11.+\'             -> Start C11 followed by 1 or more characters e.g C11_LS_100')
    else:
        ctx.obj['twinsoft_processor'].generate_tags(pattern)
        logger.info('Generate operation completed.')


@main.command()
@click.option('--tag_filter', required=False, help='Twinsoft Tag Name filter regex pattern Default: .+')
@click.option('--group_filter', required=True, help='Twinsoft Group  regex pattern')
@click.option('--dest', required=False, help='Destination Folder in Twinsoft. If not provided, mirror group_filter pattern')
@click.option('--loop', required=True, help='Loop number to ensure tags ang groups are unique')
@click.option('--offset', required=True, type=int, help='Address Offset to shift tags into')
@click.option('--replace_pattern', required=False, help='Replacement filter regex pattern. Default: \\d')
@click.pass_context
def clone(ctx, tag_filter, group_filter, dest, loop, offset, replace_pattern):
    '''
    Clone folder from twinsoft export XML file

    Most cases Tags are of the form LT_001, TIC_001_SP where 001 is the loop number or some other grouping

    Examples:

        Standard Tag Pattern:

            From LT_101 LT_115, TI_102, TIC_103 are in group CHAMBER 1 starting at adress 1000

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500

        Altenative Tag Pattern:

            From C11_LT_101 C11_LT_115, C11_TI_102, C11_ TIC_103 are in group CHAMBER 1 starting at adress 1000

            To C12_LT_101 C12_LT_115, C12_TI_102, C12_ TIC_103 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 12  --offset 1500 --replace_pattern \d{1,2}

        Alternative Tag Pattern:    

            From LT_101 LT_115, TI_102, TIC_103 are in group CHAMBER 1 starting at adress 1000

            To LT_201 LT_215, TI_202, TIC_203 are in group SPECIAL FOLDER with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500 --dest "SPECIAL FOLDER"


    '''
    spattern = tag_filter
    sreplace_pattern = replace_pattern

    if tag_filter is None:
        spattern = "^.+\d.+"

    if replace_pattern is None:
        sreplace_pattern = "\d"

    ctx.obj['twinsoft_processor'].clone(
        spattern, group_filter, dest, offset, loop, sreplace_pattern)
    logger.info('Clone operation completed.')


@main.command()
@click.argument('item')
@click.pass_context
def tabulate(ctx, item):
    '''
    Tabulate input data and copy results to clipboard

    Arguments:
                    item: xmlsummary | tags | map | template
    '''
    x = None
    log_message = None
    if item == 'xmlsummary':

        ctx.obj['twinsoft_processor'].load_twinsoft_xml()
        x = ctx.obj['twinsoft_processor'].get_twinsoft_export_summary()
        log_message = 'Summarizing Twinsoft XML ' + ctx.obj['twinsoft_export']

    elif item == 'map':
        x = ctx.obj['excel_processor'].memory_map_df
        log_message = 'Excel sheet ' + ExcelProcessor.EXCEL_MEMORY_MAP_SHEET + \
            ' in file ' + ctx.obj['excel']

    elif item == 'template':
        x = ctx.obj['excel_processor'].template_df
        log_message = 'Excel sheet ' + ExcelProcessor.EXCEL_TEMPLATE + \
            ' in file ' + ctx.obj['excel']

    elif item == 'tags':
        x = ctx.obj['excel_processor'].tags_df
        log_message = 'Excel sheet ' + ExcelProcessor.EXCEL_TAG_SHEET + \
            ' in file ' + ctx.obj['excel']

    if x is not None:
        logger.info(log_message + '\n{}'.format(x))
        x.to_clipboard()
    else:
        logger.error('Invalid item \'' + item + '\' tabulate command')


def start():
    main(obj={})


if __name__ == '__main__':
    # main("","","","")
    try:

        start()

    except ExcelProcessorError as e:
        logger.error(str(e))
    except TwinsoftError as e:
        logger.error('{} Error Code: {}'. format(
            str(e), str(e.extended_error)))
