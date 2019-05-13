# スクリプトで共通で使いそうなコマンド行オプション処理

import argparse
import logging

# 共通のオプションと固有のオプションの書き分け、
# どういう風に書くのが良いのか。とりあえずシンプルに

def args_setup(desc,defaults):
    parser = argparse.ArgumentParser(description = desc,
                                     formatter_class=argparse.RawTextHelpFormatter
    )
    dgroup = parser.add_mutually_exclusive_group()
    dgroup.add_argument("-v", "--verbose",
                        default=0, action="count",
                        help="Increase verbosity"
    )
    dgroup.add_argument("-q", "--quiet",
                        action="store_true",
                        help="Run quietly"
    )
    parser.add_argument("--icons-path",
                        help="Icons(series of images) directory. Default relative to template's './icons'")
    parser.add_argument("--pdf-export",
                        help="PDF export directory, Default './pdfout'")
    parser.add_argument("--loport",
                        help="UNO-URL, to connect to LibreOffice\n  ex. {}".format(defaults['use_loport']))
    parser.add_argument("--experimental-spillfix",
                        type=int, choices=[1,2],
                        help="Experimental spillout fix function enabled.\n" \
                        + "When 2 specified, PDF are not written which failed to fix")
    parser.add_argument("-t", "--odt",
                        help=".odt file(LoWriter document) to be used as Template")
    parser.add_argument("csv",
                        help="CSV records to be inserted into template")
    return parser

def args_parse(parser,defvalues):
    args = parser.parse_args()

    if args.quiet:
        defvalues['loglevel'] = logging.ERROR
    else:
        if args.verbose >= 2:
            defvalues['loglevel'] = logging.DEBUG
        elif args.verbose >=1:
            defvalues['loglevel'] = logging.INFO

    if args.icons_path:
        defvalues['use_iconsdir'] = args.icons_path
    if args.pdf_export:
        defvalues['use_pdfexport'] = args.pdf_export
    if args.loport:
        defvalues['use_loport'] = args.LOPORT
    if args.odt:
        defvalues['template_file'] = args.odt
    if args.csv:
        defvalues['csv_file'] = args.csv
    if args.experimental_spillfix:
        defvalues['spillfix'] = args.experimental_spillfix

    return args
