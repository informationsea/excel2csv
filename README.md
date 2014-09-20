excel2csv
=========

xls/xlsx/csv/tsv converter

License
-------

    excel2csv  xls/xlsx/csv/tsv converter
    Copyright (C) 2014 Yasunobu OKAMURA

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.


### Binary distribution license

The binary distribution of excel2csv includes following open source
software. They are licensed by Apache License.

* [Apache POI] - the Java API for Microsoft Documents
* [Open CSV] - a very simple csv (comma-separated values) parser
  library for Java
* [Apache Commons CLI] - API for parsing command line options passed to programs

[Apache POI]: http://poi.apache.org
[Open CSV]: http://opencsv.sourceforge.net
[Apache Commons CLI]: http://commons.apache.org/proper/commons-cli/


Install
-------

1. Copy bin/excel2csv to /usr/local/bin
2. Copy lib/excel2csv.jar to /usr/local/lib

Usage
-----

    usage: excel2csv [options] [INPUT] [OUTPUT]
     -A         file type of output file will detect automatically (default)
     -a         file type of input file will detect automatically (default)
     -F         Overwrite sheet if exists
     -h         show help
     -i <arg>   Sheet index of input file (xls/xlsx only / default: 0)
     -s <arg>   Sheet name of input file (xls/xlsx only)
     -S <arg>   Sheet name candidate of output file (xls/xlsx only)

The file format is automatically detected by suffix. If an output
xls/xlsx file is exists, excel2csv do not overwrite file. In this case
excel2csv add a sheet to the file. By default, a name the of sheet is
suggested from a input file name.

### Convert xls to csv

    excel2csv iris.xls iris.csv
    
### Convert csv to xlsx and set sheet name as "iris"

    excel2csv -S iris iris.csv iris.xlsx

### Add tsv data to xlsx and set sheet name as "foo"

    excel2csv -S foo foo.txt iris.xlsx

How to build
------------

    git clone https://github.com/informationsea/excel2csv.git
    cd excel2csv
    ./build.sh



