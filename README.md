[![Kwalitee status](https://cpants.cpanauthors.org/dist/Spreadsheet-SimpleExcel.png)](https://cpants.cpanauthors.org/dist/Spreadsheet-SimpleExcel)
[![GitHub issues](https://img.shields.io/github/issues/reneeb/Spreadsheet-SimpleExcel.svg)](https://github.com/reneeb/Spreadsheet-SimpleExcel/issues)
[![CPAN Cover Status](https://cpancoverbadge.perl-services.de/Spreadsheet-SimpleExcel-1.92)](https://cpancoverbadge.perl-services.de/Spreadsheet-SimpleExcel-1.92)
[![Cpan license](https://img.shields.io/cpan/l/Spreadsheet-SimpleExcel.svg)](https://metacpan.org/release/Spreadsheet-SimpleExcel)

# NAME

Spreadsheet::SimpleExcel - Create Excel files with Perl

# VERSION

version 1.92

# SYNOPSIS

```perl
use Spreadsheet::SimpleExcel;

binmode(\*STDOUT);
# data for spreadsheet
my @header = qw(Header1 Header2);
my @data   = (['Row1Col1', 'Row1Col2'],
              ['Row2Col1', 'Row2Col2']);

# create a new instance
my $excel = Spreadsheet::SimpleExcel->new();

# add worksheets
$excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
$excel->add_worksheet('Second Worksheet',{-data => \@data});
$excel->add_worksheet('Test');

# add a row into the middle
$excel->add_row_at('Name of Worksheet',1,[qw/new row/]);

# sort data of worksheet - ASC or DESC
$excel->sort_data('Name of Worksheet',0,'DESC');

# remove a worksheet
$excel->del_worksheet('Test');

# sort worksheets
$excel->sort_worksheets('DESC');

# create the spreadsheet
$excel->output();

# print sheet-names
print join(", ",$excel->sheets()),"\n";

# get the result as a string
my $spreadsheet = $excel->output_as_string();

# print result into a file and handle error
$excel->output_to_file("my_excel.xls") or die $excel->errstr();
$excel->output_to_file("my_excel2.xls",45000) or die $excel->errstr();

## or

# data
my @data2  = (['Row1Col1', 'Row1Col2'],
              ['Row2Col1', 'Row2Col2']);

my $worksheet = ['NAME',{-data => \@data2}];
# create a new instance
my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

# add headers to 'NAME'
$excel2->set_headers('NAME',[qw/this is a test/]);
# append data to 'NAME'
$excel2->add_row('NAME',[qw/new row/]);

$excel2->output();

$excel2->output_to_XML('test.xml');

## create XLSX
my $worksheet3 = [ 'NAME', { -data => \@data } ];
my $file3      = 'test.xlsx';

# create a new instance
my $excel3 = Spreadsheet::SimpleExcel->new(
    -worksheets => [$worksheet3],
    -filename   => $file3,
    -format     => 'xlsx',
);

# add headers to 'NAME'
$excel3->set_headers('NAME',[qw/this is a test/]);

$excel3->output_to_file();
```

# DESCRIPTION

Spreadsheet::SimpleExcel simplifies the creation of excel-files in the web. It does
provide simple cell-formats, but only three types of formats (to keep the module simple).

# METHODS

Added in version 1.4:

If you want a method to do the functionality for the last inserted worksheet
(current sheet), you don't have to pass the title as a parameter for the method.

So now you can do something like this:

```
$excel->add_worksheet("Test");
$excel->add_row(\@data);
$excel->sort_date($column_idx);
```

This leads to more usability.

## new

```perl
# create a new instance
my $excel = Spreadsheet::SimpleExcel->new();

# or

my $worksheet = ['NAME',{-data => ['This','is','an','Test']}];
my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

# to create a file
my $filename = 'test.xls';
my $excel = Spreadsheet::SimpleExcel->new(-filename => $filename);

#if a file > 7 MB should be created
$excel = Spreadsheet::SimpleExcel->new(-big => 1);
```

If -big is set to true, Spreadsheet::WriteExcel::Big is required!

## add\_worksheet

```perl
# add worksheets
$excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
$excel->add_worksheet('Second Worksheet',{-data => \@data});
$excel->add_worksheet('Test');
```

The first parameter of this method is the name of the worksheet and the second one is
a hash with (optional) information about the headlines and the data.
No duplicate worksheets allowed.

## del\_worksheet

```
# remove a worksheet
$excel->del_worksheet('Test');
```

Deletes all worksheets named like the first parameter

## add\_row

```
# append data to 'NAME'
$excel->add_row('NAME',[qw/new row/]);
```

Adds a new row to the worksheet named 'NAME'

## add\_row\_at

```
# add a row into the middle
$excel->add_row_at('Name of Worksheet',1,[qw/new row/]);
```

This method inserts a row into the existing data

## sort\_data

```
# sort data of worksheet - ASC or DESC
$excel->sort_data('Name of Worksheet',0,'DESC');
```

sort\_data sorts the rows. All sorts for one worksheet are combined, so 

```
$excel->sort_data('Name of Worksheet',0,'DESC');
$excel->sort_data('Name of Worksheet',1,'ASC');
```

will sort the column 0 first and then (within this sorted data) the
column 1.

## reset\_sort

```
$excel->reset_sort('Name of Worksheet');
```

The data won't be sorted, the data are in original order instead.

## set\_headers

```
# add headers to 'NAME'
$excel->set_headers('NAME',[qw/this is a test/]);
```

set the headers for the worksheet named 'NAME'

## errstr

returns error message.

## sort\_worksheets

```
# sort worksheets
$excel->sort_worksheets('DESC');
```

sorts the worksheets in DESCending or ASCending order.

## output

```
$excel2->output();
```

prints the worksheet to the STDOUT and prints the Mime-type 'application/vnd.ms-excel'.

## output\_as\_string

```perl
# get the result as a string
my $spreadsheet = $excel->output_as_string();
```

returns a string that contains the data in excel-format

## output\_to\_file

```perl
# print result into a file [output_to_file(<filename>,<lines>)]
$excel->output_to_file("my_excel.xls");
$excel->output_to_file("my_excel2.xls",45000) or die $excel->errstr();
```

prints the data into a file.
The data will be printed into more worksheets, if the number of rows is greater than &lt;lines> (default 32000).

## output\_to\_XML

```
$excel2->output_to_XML('test.xml');
```

prints the data into a XML file.

## sheets

```
$ref = $excel->sheets();
@names = $excel->sheets();
```

In listcontext this subroutines returns a list of the names of sheets that are in $excel, in
scalar context it returns a reference on an Array.

## set\_headers\_format

```
# set formats for headers of 'NAME'
# first col 'string', second col 'number', third col default format, fourth col 'number'
$excel2->set_headers_format('NAME',['s','n',undef,'n']);
```

sets the headers formats for a specified worksheet. If formats are commited, the default
format is set. Default format is set by Spreadsheet::WriteExcel

## set\_data\_format

```
# set formats for headers of 'NAME'
# first col 'string', second col 'number', third col default format, fourth col 'number'
$excel2->set_data_format('NAME',['s','n',undef,'n']);
```

sets the data formats for a specified worksheet. If formats are commited, the default
format is set. Default format is set by Spreadsheet::WriteExcel

## current\_sheet

```
$excel->add_worksheet('Testtitle');
print $excel->current_sheet;
```

returns the title of the current worksheet.

# EXAMPLES

## PRINT ON STDOUT

```perl
#! /usr/bin/perl

use strict;
use warnings;
use Spreadsheet::SimpleExcel;

binmode(\*STDOUT);
# data for spreadsheet
my @header = qw(Header1 Header2);
my @data   = (['Row1Col1', 'Row1Col2'],
              ['Row2Col1', 'Row2Col2']);

# create a new instance
my $excel = Spreadsheet::SimpleExcel->new();

# add worksheets
$excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
$excel->add_worksheet('Second Worksheet',{-data => \@data});
$excel->add_worksheet('Test');

# add a row into the middle
$excel->add_row_at('Name of Worksheet',1,[qw/new row/]);

# sort data of worksheet - ASC or DESC
$excel->sort_data('Name of Worksheet',0,'DESC');

# remove a worksheet
$excel->del_worksheet('Test');

# create the spreadsheet
$excel->output();
```

## RECEIVE DATA AS A SCALAR

```perl
#!/usr/bin/perl

use strict;
use warnings;
use Spreadsheet::SimpleExcel;

# data
my @data2  = (['Row1Col1', 'Row1Col2'],
              ['Row2Col1', 'Row2Col2']);

my $worksheet = ['NAME',{-data => \@data2}];
# create a new instance
my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

# add headers to 'NAME'
$excel2->set_headers('NAME',[qw/this is a test/]);
# append data to 'NAME'
$excel2->add_row('NAME',[qw/new row/]);

# receive as string
my $string = $excel2->output_as_string();
```

## PRINT INTO FILE

```perl
#! /usr/bin/perl

use strict;
use warnings;
use Spreadsheet::SimpleExcel;

# data
my @data2  = (['Row1Col1', 'Row1Col2'],
              ['Row2Col1', 'Row2Col2']);

my $worksheet = ['NAME',{-data => \@data2}];
# create a new instance
my $excel2    = Spreadsheet::SimpleExcel->new(-worksheets => [$worksheet]);

# add headers to 'NAME'
$excel2->set_headers('NAME',[qw/this is a test/]);
# append data to 'NAME'
$excel2->add_row('NAME',[qw/new row/]);

# print into file
$excel2->output_to_file("my_excel.xls");
```

## PRINT INTO FILE (break worksheets)

```perl
#! /usr/bin/perl

use strict;
use warnings;
use Spreadsheet::SimpleExcel;

# create a new instance
my $excel    = Spreadsheet::SimpleExcel->new();

my @header = qw(Header1 Header2);
my @data   = (['Row1Col1', 'Row1Col2'],
              ['Row2Col1', 'Row2Col2']);
for(0..70000){
  push(@data,[qw/1 2 4 6 8/]);
}
# add worksheets
$excel->add_worksheet('Name of Worksheet',{-headers => \@header, -data => \@data});
$excel->add_row('Name of Worksheet',[qw/1 2 3 4 5/]);

# print into file
$excel->output_to_file("my_excel.xls",10000);
```

# DEPENDENCIES

This module requires Spreadsheet::WriteExcel and IO::Scalar

# SEE ALSO

Spreadsheet::WriteExcel

IO::Scalar

IO::File

XML::Writer



# Development

The distribution is contained in a Git repository, so simply clone the
repository

```
$ git clone git://github.com/reneeb/Spreadsheet-SimpleExcel.git
```

and change into the newly-created directory.

```
$ cd Spreadsheet-SimpleExcel
```

The project uses [`Dist::Zilla`](https://metacpan.org/pod/Dist::Zilla) to
build the distribution, hence this will need to be installed before
continuing:

```
$ cpanm Dist::Zilla
```

To install the required prequisite packages, run the following set of
commands:

```
$ dzil authordeps --missing | cpanm
$ dzil listdeps --author --missing | cpanm
```

The distribution can be tested like so:

```
$ dzil test
```

To run the full set of tests (including author and release-process tests),
add the `--author` and `--release` options:

```
$ dzil test --author --release
```

# AUTHOR

Renee Baecker <reneeb@cpan.org>

# COPYRIGHT AND LICENSE

This software is Copyright (c) 2015 by Renee Baecker.

This is free software, licensed under:

```
The Artistic License 2.0 (GPL Compatible)
```
