#!perl -w
use strict;

#use strict "refs", "subs";
use warnings;
use Carp 'verbose';
$SIG{__DIE__} = sub { Carp::confess(@_) };
use Data::Dumper;
use Storable;
use Unicode::Collate;
use Spreadsheet::Read;
my ( $dir, $book, $sheet, $cell, $row, $col, $book_copy );

#    use Config;    # OS dependent configuration section
if    ( $^O =~ /^MSWin32/xi ) { $dir = '/home/westmj/Downloads/'; }
elsif ( $^O =~ /^linux/xi )   { $dir = '/home/westmj/Downloads/'; }
elsif ( $^O =~ /^darwin/xi )  { $dir = '/Users/Headofschool/Downloads/'; }

#        else {
#            if ( -f $arg ) { $excel_file = $arg }
#            if ( -d $arg ) { $output_dir = abs_path($arg) }
#        }
$book = ReadData( "$dir" . 'MFSRoster2018-2019.xlsx' );    # MFSRoster2018-2019
store \$book, $dir . 'store';
$book_copy = retrieve( $dir . 'store' );
print "\$^O = '$^O'.  Done.\n";
