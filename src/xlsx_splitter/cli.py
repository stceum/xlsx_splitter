import argparse
import sys
from .splitter import ExcelSplitter


def run():
    import sys
    class CustomParser(argparse.ArgumentParser):
        def error(self, message):
            sys.stderr.write('error: %s\n' % message)
            self.print_help()
            sys.exit(2)

    parser = CustomParser()
    parser.add_argument("-f", "--file", required=True)
    parser.add_argument("-o", "--output", default="")
    parser.add_argument("-n", "--number", required=True, type=int)
    parser.add_argument("-s", "--sheet-name", default="Sheet1")
    parser.add_argument("--save-in-one-file", action="store_true")
    args = parser.parse_args()

    if args.output == "":
        args.output = args.file.rsplit('.', 1)[0]+".split.xlsx"
    ExcelSplitter(args.file).split_by_row_in_average(args.sheet_name, args.number, args.output, save_in_one_file=args.save_in_one_file)

