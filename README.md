# A XLSX Splitter

The excel editors always freeze when processing large excel files. Here come the tool to split a large excel into small ones.

This tool is build with python pacakge `openpyxl`.

## Usage

```shell
usage: xlsx-split [-h] -f FILE [-o OUTPUT] -n NUMBER [--save-in-one-file]

options:
  -h, --help            show this help message and exit
  -f FILE, --file FILE
  -o OUTPUT, --output OUTPUT
  -n NUMBER, --number NUMBER
  -s SHEET_NAME, --sheet-name SHEET_NAME  
  --save-in-one-file
```

- spilt file `/path/to/your/input.xlsx` into `N` files equally, to the directory `/output/dir`, like `/output/dir/split1.xlsx`, `/output/dir/split2.xlsx` ...

  ```shell
  xlsx-split -f /path/to/your/input.xlsx -o /output/dir -n N
  ```

- spilt file `/path/to/your/input.xlsx` into `N` files equally, to the file `/path/to/output.xlsx`, saving in one file with `N` sheets

  ```shell
  xlsx-split -f /path/to/your/input.xlsx -o /path/to/output.xlsx -n N --save-in-one-file
  ```

## TODOs

- [x] Split into Target Number of Files
- [ ] Rich Format Support
