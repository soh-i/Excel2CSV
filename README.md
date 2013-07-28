Excel2CSV
=========

Yet another xls/xlsx file conversion tool.
The simple script converts .xls/.xlsx file to csv, properly converting by the given file suffix (.xls or .xlsx).

## Installation
* The script depends the following CPAN modules,
* `cpanm Spreadsheet::ParseExcel Spreadsheet::XLSX`

## Usage
* `./xls2csv --excel data.xls --sheet Sheet1`
* `./xls2csv --excel accounts.xlsx --sheet Name1`


### Required arguments
* `-e,  --excel`     Given a .xls or .xlsx file.       [Required]
* `-s,  --sheet`     Given a sheet name of the file.   [Required]

### Options
* `-h,  --help `     Show help messages.

## Contributing
* Fork it
* Create your feature branch (`git checkout -b my-new-feature`)
* Commit your changes (`git commit -am 'Add some feature'`)
* Push to the branch (`git push origin my-new-feature`)
* Create new Pull Request
