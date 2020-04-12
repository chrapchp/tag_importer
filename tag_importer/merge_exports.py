import pandas as pd
import argparse
import sys
import logging
import xml.etree.ElementTree as et
import glob
import xml.dom.minidom
'''
merge one or more xml files 
'''


def merge(pattern, outfile):
    xml_files = glob.glob(pattern)
    xml_main = None
    for xml_file in xml_files:
        print("Processing {}...".format(xml_file))
        tree = et.parse(xml_file)
        root = tree.getroot()
        if xml_main is None:
            xml_main = tree
        else:
            for tag in root.findall('Tag'):
                xml_main.getroot().append(tag)
    if xml_main is not None:
        print("Writing to {}".format(outfile))
        xml_main.write(outfile)
    else:
        print("No XML files found in pattern {}".format(pattern))


if __name__ == '__main__':

    ap = argparse.ArgumentParser()
    ap.add_argument("-p", "--pattern", required=True,
                    help="file pattern to merge. e.g p*.xml ")
    ap.add_argument("-o", "--outfile", required=True, help="output file name")

    args = vars(ap.parse_args())
    files = glob.glob(sys.argv[1])

    merge(args['pattern'], args['outfile'])
