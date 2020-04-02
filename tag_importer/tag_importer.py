
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
            format='%(asctime)s - %(message)s', level=logging.INFO)
        # logging.basicConfig(level=logging.INFO)


@main.command()
@click.option('--pattern', required=True, help='Generate Tags for a given TAG_PATTERN defined in excel TAGS tab')
@click.pass_context
def generate(ctx, pattern):
    '''
    Generate tags using pattern defined in XL
    '''
    ctx.obj['twinsoft_processor'].generate_tags(pattern)


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
        logger.error('{} Error Code: {}'. format(str(e), str(e.extended_error)))
