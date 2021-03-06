
import pandas as pd
#import argparse
#import sys
import logging
import os
import sys
import verboselogs
import click


from twinsoft import TwinsoftProcessor
from twinsoft import TwinsoftError
from helpers import ExcelProcessor
from helpers import ExcelProcessorError


logger = verboselogs.VerboseLogger(__name__) #logger = logging.getLogger(__name__)


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
        FORMAT = "[%(asctime)s %(levelname)s-%(filename)s %(funcName)s() ] %(message)s"
        logging.basicConfig(level=verboselogs.VERBOSE, format=FORMAT)
    else:
        logging.basicConfig(
            format='%(asctime)s %(levelname)s- %(message)s', level=logging.INFO)

@main.command()
@click.option('--pattern', required=True, help='Generate Tags for a given TAG_PATTERN defined in excel TAGS tab')
@click.option('--tag_type', required=False, help='local | remote', default='local')
@click.option('--ignore_map_errors/--no-ignore_map_errors', required=False, help='ignore mapping errors', default=False)
@click.pass_context
def generate(ctx, pattern,tag_type,ignore_map_errors):
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
    elif tag_type=='local':
        ctx.obj['twinsoft_processor'].generate_tags(pattern,ignore_map_errors)
        logger.info('Generate tags operation completed.')
    elif tag_type=='remote':
        ctx.obj['twinsoft_processor'].generate_remote_tags(pattern)
        #logger.info('Generate remote tags operation completed.')
    else:
        logger.error('Invalid tag_type: ' + tag_type + ' for generate command. local | remote')

@main.command()
@click.option('--tag_filter', required=False, default=r"^.+\d.+", help='Twinsoft Tag Name filter regex pattern Default: .+')
@click.option('--group_filter', required=True, help='Twinsoft Group  regex pattern')
@click.option('--recurse/--no-recurse', default=True, help='Recurse Folder e.g. CHAMBER 1 and CHAMBER 1/SOFTS. default:--recurse')
@click.pass_context
def create(ctx, tag_filter, group_filter,recurse):
    sgroup_filter = group_filter
    if not recurse:
        sgroup_filter = "^"+group_filter+"$"
    ctx.obj['twinsoft_processor'].create( tag_filter, sgroup_filter)
    logger.info('Create operation completed.')

@main.command()
@click.option('--tag_filter', required=False, default=r"^.+\d.+", help='Twinsoft Tag Name filter regex pattern Default: .+')
@click.option('--group_filter', required=True, help='Twinsoft Group regex pattern')
@click.option('--dest', required=False, help='Destination Folder in Twinsoft. If not provided, mirror group_filter pattern')
@click.option('--loop', required=True, help='Loop number to ensure tags and groups are unique')
@click.option('--offset', required=True, type=int, help='Address Offset to shift tags into')
@click.option('--replace_pattern', required=False, help='Replacement filter regex pattern for tags and groups. Default: \\d')
@click.option('--recurse/--no-recurse', default=True, help='Recurse Folder e.g. CHAMBER 1 and CHAMBER 1/SOFTS. default:--recurse')
@click.option('--blind_validation/--no-blind_validation', default=False, help='ignore missing memory map entries')
@click.option('--group_find', required=False, default=None, help='Find group_find and replace with group_replace')
@click.option('--group_replace', required=False, default=None, help='Find group_find and replace with group_replace')
@click.option('--ignore_map_errors/--no-ignore_map_errors', required=False, default=False, help='ignore overlaps in memory between binary and analogs')
@click.pass_context
def clone(ctx, tag_filter, group_filter, dest, loop, offset, replace_pattern,recurse,blind_validation,group_find,group_replace,ignore_map_errors):
    '''
    Clone folder from twinsoft export XML file

    Most cases Tags are of the form LT_001, TIC_001_SP where 001 is the loop number or some other grouping

    Examples:

        Standard Tag Pattern:

            From LT_101 LT_115, TI_102, TIC_103 are in group CHAMBER 1 and CHAMBER 1\SOFTS containts LT_101_SP starting at adress 1000

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 and LT_201_SP in CHAMBER 2\SOFTS with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500

        Altenative Tag Pattern:

            From C11_LT_101 C11_LT_115, C11_TI_102, C11_ TIC_103 are in group CHAMBER 1 starting at adress 1000

            To C12_LT_101 C12_LT_115, C12_TI_102, C12_ TIC_103 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 12  --offset 1500 --replace_pattern \d{1,2}

        Alternative Tag Pattern:    

            From LT_101 LT_115, TI_102, TIC_103 are in group CHAMBER 1 starting at adress 1000

            To LT_201 LT_215, TI_202, TIC_203 are in group SPECIAL FOLDER with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500 --dest "SPECIAL FOLDER"
        
        Altenative Tag Pattern:
            
            From LT_101 LT_115, TI_102, TIC_103 are in group CHAMBER 1 and CHAMBER 1\SOFTS containts LT_101_SP starting at adress 1000

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 AND don't clone subfolders with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500  --no-recurse      

        Altenative Tag Pattern:
            
            Clone all tags in group CHAMBER 1 with tags starting at address 1000 

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500  --tag_filter AT.+

        Altenative Tag Pattern:
            
            Clone all tags in group CHAMBER 1 with tags starting with at address 1000 and igonore memory map conflicts between binaries and analogs

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500  --tag_filter AT.+  --ignore_map_errors          


        Altenative Tag Pattern:
            
            Clone all tags in group CHAMBER 1 with tags starting with at address 1000 and igonore missing memory map entries

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500  --tag_filter AT.+  --blind_validation       

        Altenative Tag Pattern:
            
            Clone all tags in group CHAMBER 1 with tags starting with at address 1000 and force setting group name to SPECIAL

            To LT_201 LT_215, TI_202, TIC_203 are in group CHAMBER 2 with address starting at 2500

            pass options --group_filter "CHAMBER 1" --loop 2  --offset 1500  --group_find "CHAMBER 1" --dest "SPECIAL"             


    '''

    sreplace_pattern = replace_pattern
    sgroup_filter = group_filter
    
    if not recurse:
        sgroup_filter = "^"+group_filter+"$"


    if group_find is not None and group_replace is None:
        raise ExcelProcessorError("option --group_replace required")

    if replace_pattern is None:
        sreplace_pattern = "\d"

    ctx.obj['twinsoft_processor'].clone(
        tag_filter, sgroup_filter, dest, offset, loop, sreplace_pattern,blind_validation,group_find, group_replace,ignore_map_errors)
    logger.info('Clone operation completed.')


@main.command()
@click.argument('item')
@click.option('--mapped/--no-mapped', default=False)
@click.option('--root_tags/--no_root_tags', default=False)
@click.pass_context
def tabulate(ctx, item,mapped,root_tags):
    '''
    Tabulate input data and copy results to clipboard

    Arguments:
                    item: xmlsummary | tags | map | template

    --mapped  will cause xmlmsummary will be formated as an MEMORY_MAP format for Excel
    '''
    x = None
    log_message = None
    if item == 'xmlsummary':

        ctx.obj['twinsoft_processor'].load_twinsoft_xml(validate=False)

       
        if mapped:
            log_message = 'Summarizing Twinsoft XML ' + ctx.obj['twinsoft_export'] + ' as MEMORY_MAP REMOVE INDEX'
        else:
            log_message = 'Summarizing Twinsoft XML ' + ctx.obj['twinsoft_export']

        x = ctx.obj['twinsoft_processor'].get_twinsoft_export_summary(mapped,root_tags)


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
        if os.name != 'POSIX':
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
        logger.error('{} Error Code: {} '. format(
            str(e), str(e.extended_error)))
    #except pd.io.clipboard.PyperclipException as e:
    #    logger.warning("Clipboard not install in POSIX build. to_clipboard()" )

