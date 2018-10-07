#! perl -w
use strict;
use warnings;
use Carp 'verbose';
local $SIG{__DIE__} = sub { Carp::confess(@_) };
use Data::Dumper;
use Storable;
use Spreadsheet::WriteExcel;

use Getopt::Long;
my $dir;
if    ( $^O =~ /^MSWin32/xi ) { $dir = '/home/westmj/Downloads/'; }
elsif ( $^O =~ /^linux/xi )   { $dir = '/home/westmj/Downloads/'; }
elsif ( $^O =~ /^darwin/xi )  { $dir = '/Users/Headofschool/Downloads/'; }
my $book_store    = 'store';
my $workbook_name = 'directory.xls';
GetOptions(
    "dir=s"        => \$dir,
    "book_store=s" => \$book_store,
    "workbook=s"   => \$workbook_name
) or die("Error in command line arguments\n");

# print "dir '$dir' book_store '$book_store' workbook_name '$workbook_name'\n"; exit;

my (
    $fh,     $fh_string,   $sheet,       $father,
    $mother, $father_name, $mother_name, %family,
    @staff,  @committee,   @volunteers,  $table,
    $row,    $guardians,   $guardians_name,
);
my @grade = (
    '',
    '',
    '',
    'PreKinder',
    'Kinder',
    'Preparatoria',
    '1er Grado | 1st Grade',
    '2do Grado | 2nd Grade',
    '3er Grado | 3rd Grade',
    '4to Grado | 4th Grade',
    '5to Grado | 5th Grade',
    '6to Grado | 6th Grade',
    '7mo Grado | 7th Grade',
    '8vo Grado | 8th Grade',
    '9no Grado | 9th Grade',
    '10mo Grado | 10th Grade',
    '11mo Grado | 11th Grade',
    '12mo Grado | 12h Grade',
    'Estudiante Año Sábatico | GAP Year'
);

my $book_copy       = retrieve( $dir . $book_store );
my $workbook        = Spreadsheet::WriteExcel->new( $dir . $workbook_name );
my $worksheet       = $workbook->add_worksheet('staff');
my $worksheet_index = $workbook->add_worksheet('index');

#staff
foreach my $table ( 1 .. 2 ) {
    $sheet = $$book_copy->[$table]{cell};
    foreach my $row ( 2 .. scalar( @{$sheet} ) ) {
        if ( !$$book_copy->[$table]{cell}[1][$row] ) { next }
        if ( !$$book_copy->[$table]{cell}[1][$row] or
         !$$book_copy->[$table]{cell}[2][$row] or
         !$$book_copy->[$table]{cell}[16][$row] or
        !$$book_copy->[$table]{cell}[7][$row] or
        !$$book_copy->[$table]{cell}[4][$row]        ) {
 print "Something is empty: In table '$table' grade $grade[$table] row '$row'  
               first name '$$book_copy->[$table]{cell}[1][$row]' 
               last name '$$book_copy->[$table]{cell}[2][$row]' 
               phone '$$book_copy->[$table]{cell}[7][$row]'
               email '$$book_copy->[$table]{cell}[4][$row]' \n"; 
        }
        push @staff, $$book_copy->[$table]{cell}[1][$row] . ' ' .   # first name
          $$book_copy->[$table]{cell}[2][$row] . ' (' .  # last name
          $$book_copy->[$table]{cell}[16][$row] . ')  '
          .                                              # annotation / position
          $$book_copy->[$table]{cell}[7][$row] . ' ' .   # phone
          $$book_copy->[$table]{cell}[4][$row] . "\n";   # email
    }
}
open $fh, ">:encoding(UTF-8)", $dir . "mfs_staff.txt" or croak $!;
print $fh sort(@staff);
close $fh;

#excel staff
$row = 0;
foreach my $line ( sort @staff ) {
    chomp $line;
    $worksheet->write( $row++, 0, $line );
}

#volunteers
$worksheet = $workbook->add_worksheet('volunteers');
foreach my $table ( 20 .. 20 ) {
    $sheet = $$book_copy->[$table]{cell};
    foreach my $row ( 2 .. scalar( @{$sheet} ) ) {
        if ( !$$book_copy->[$table]{cell}[1][$row] ) { next }
        if ( !$$book_copy->[$table]{cell}[16][$row] ) {
            next;
        }    # only process volunteers with something in annotation
        push @volunteers,
          $$book_copy->[$table]{cell}[1][$row] . ' ' .   # first name
          $$book_copy->[$table]{cell}[2][$row] . ' (' .  # last name
          $$book_copy->[$table]{cell}[16][$row] . ')  '
          .                                              # annotation / position
          $$book_copy->[$table]{cell}[7][$row] . ' ' .   # phone
          $$book_copy->[$table]{cell}[4][$row] . "\n";   # email
    }
}
open $fh, ">:encoding(UTF-8)", $dir . "mfs_volunteers.txt" or croak $!;
print $fh sort(@volunteers);
close $fh;

#excel volunteers
$row = 0;
foreach my $line ( sort @volunteers ) {
    chomp $line;
    $worksheet->write( $row++, 0, $line );
}

#committee
$worksheet = $workbook->add_worksheet('committee');
foreach my $table ( 21 .. 21 ) {
    $sheet = $$book_copy->[$table]{cell};
    foreach my $row ( 2 .. scalar( @{$sheet} ) ) {
        if ( !$$book_copy->[$table]{cell}[1][$row] ) { next }
        push @committee,
          $$book_copy->[$table]{cell}[1][$row] . ' ' .   # first name
          $$book_copy->[$table]{cell}[2][$row] . ' (' .  # last name
          $$book_copy->[$table]{cell}[16][$row] . ')  '
          .                                              # annotation / position
          $$book_copy->[$table]{cell}[7][$row] . ' ' .   # phone
          $$book_copy->[$table]{cell}[4][$row] . "\n";   # email
    }
}
open $fh, ">:encoding(UTF-8)", $dir . "mfs_committee.txt" or croak $!;
print $fh sort(@committee);
close $fh;

#excel committee
$row = 0;
foreach my $line ( sort @committee ) {
    chomp $line;
    $worksheet->write( $row++, 0, $line );
}

# family
$fh_string = '';
foreach my $table ( 3 .. 17 ) {
    $worksheet = $workbook->add_worksheet( $grade[$table] );    # name the table
    $sheet     = $$book_copy->[$table]{cell};                   # get the sheet
    $fh_string = $fh_string . "$grade[$table]\n";

    foreach my $row ( 2 .. scalar( @{$sheet} ) ) {
        if ( !${ $$book_copy->[$table]{cell}[2] }[$row] ) { next }  # empty line
        if ( !${ $$book_copy->[$table]{cell}[11] }[$row] ) {
            $father      = '';
            $father_name = '';
        }
        elsif ( ${ $$book_copy->[$table]{cell}[11] }[$row] eq '?' ) {
            $father      = '';
            $father_name = '';
        }
        else {
            $father = ${ $$book_copy->[$table]{cell}[11] }[$row];

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
 #           $father =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
# take out private information
 #           $father =~ s/\[\s*privado[^]]*\]\s*//gx;
            # take out private information
            chomp $father;

# $father =~ s/.([^\(\s]*?) \((.)(\1)\)/$2$3/;  # look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
            $father_name = $father;
            $father_name =~ s/(.*?),.*/$1/x;    # chop off phone number and all

            unless ( exists( $family{$father} ) ) { $family{$father} = '' }
            $family{$father} =
                $family{$father} . "\n"
              . '                      '
              . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
              . ${ $$book_copy->[$table]{cell}[2] }[$row] . ' ('
              . "$grade[$table])";              # add a kid
              $family{$father} =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
        }

        if ( !${ $$book_copy->[$table]{cell}[12] }[$row] ) {
            $mother      = '';
            $mother_name = '';
        }
        elsif ( ${ $$book_copy->[$table]{cell}[12] }[$row] eq '?' ) {
            $mother      = '';
            $mother_name = '';
        }
        else {
            $mother = ${ $$book_copy->[$table]{cell}[12] }[$row];

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
            $mother =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
# take out private information
            $mother =~ s/\[\s*privado[^]]*\]\s*//gx;

            # take out private information
            $mother_name = $mother;
            $mother_name =~ s/(.*?),.*/$1/x;
            unless ( exists( $family{$mother} ) ) { $family{$mother} = '' }

            $family{$mother} =
                $family{$mother} . "\n"
              . '                      '
              . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
              . ${ $$book_copy->[$table]{cell}[2] }[$row] . ' ('
              . "$grade[$table])";
        }

        # guardians

        if ( !${ $$book_copy->[$table]{cell}[13] }[$row] ) {
            $guardians      = '';
            $guardians_name = '';
        }
        elsif ( ${ $$book_copy->[$table]{cell}[13] }[$row] eq '?' ) {
            $guardians      = '';
            $guardians_name = '';
        }
        else {
            $guardians = ${ $$book_copy->[$table]{cell}[13] }[$row];

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
            $guardians =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
# take out private information
            $guardians =~ s/\[\s*privado[^]]*\]\s*//gx;

      # take out private information
      # print "table $table $grade[$table] guardian row  $row '$guardians' \n" ;
            foreach my $guardian ( split " & ", $guardians ) {
                chomp $guardian; 
                $guardian =~ s/^\s*//;
                $guardian =~ s/\s*$//;
                my $guardian_name = $guardian;
                $guardian_name =~
                  s/(.*?),.*/$1/x;    # chop off phone number and all
                unless ( exists( $family{$guardian} ) ) {
                    $family{$guardian} = '';
                }
                $family{$guardian} =
                    $family{$guardian} . "\n"
                  . '                      '
                  . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
                  . ${ $$book_copy->[$table]{cell}[2] }[$row] . ' ('
                  . "$grade[$table]) ";    # add a kid
            }
        }

        if ( !$father && !$mother ) {
            if (   ( ${ $$book_copy->[$table]{cell}[13] }[$row] eq '?' )
                or ( !${ $$book_copy->[$table]{cell}[13] }[$row] ) )
            {
                print
                  "No family at row $row in table $table '$grade[$table]' ! \n";
            }

            ( $father, $mother ) =
              ( split " & ", ${ $$book_copy->[$table]{cell}[13] }[$row] );

# print  "table $table '$grade[$table]' row $row ${$$book_copy->[$table]{cell}[13]}[$row] \n";
            $father_name = $father;
            $father_name =~ s/(.*?),.*/$1/x;    # chop off phone number and all
# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
            $father_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;

            $mother_name = $mother;
# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
            $mother_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
            $mother_name =~ s/(.*?),.*/$1/x;
            $fh_string =
                $fh_string
              . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
              . ${ $$book_copy->[$table]{cell}[2] }[$row]
              . "\t -- \t"
              . $father_name . ' & '
              . $mother_name;
            $worksheet->write_string(
                $row - 2,
                0,
                ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
                  . ${ $$book_copy->[$table]{cell}[2] }[$row]
            );
            $worksheet->write_string( $row - 2, 1,
                $father_name . ' & ' . $mother_name );
        }
        elsif ( $father && $mother ) {
            $father_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
           $mother_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
            $fh_string =
                $fh_string
              . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
              . ${ $$book_copy->[$table]{cell}[2] }[$row]
              . "\t -- \t"
              . $father_name . ' & '
              . $mother_name;
            $worksheet->write_string(
                $row - 2,
                0,
                ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
                  . ${ $$book_copy->[$table]{cell}[2] }[$row]
            );
            $worksheet->write_string( $row - 2, 1,
                $father_name . ' & ' . $mother_name );
        }
        elsif ( !$father && $mother ) {
                       $father_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
           $mother_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
            $fh_string =
                $fh_string
              . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
              . ${ $$book_copy->[$table]{cell}[2] }[$row]
              . "\t -- \t"
              . $mother_name;
            $worksheet->write_string(
                $row - 2,
                0,
                ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
                  . ${ $$book_copy->[$table]{cell}[2] }[$row]
            );
            $worksheet->write_string( $row - 2, 1, $mother_name );
        }
        else {
                       $father_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
           $mother_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
            $fh_string =
                $fh_string
              . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
              . ${ $$book_copy->[$table]{cell}[2] }[$row]
              . "\t -- \t"
              . $father_name;
        }

        # excel family

    }
    $fh_string = $fh_string . "\n\n";
}

open $fh, ">:encoding(UTF-8)", $dir . "mfs_out.txt" or croak $!;
print $fh $fh_string;
close $fh;

# index
$fh_string = '';
$row       = 0;
foreach my $key ( sort keys(%family) ) {
    my $altered = $key;

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
    $altered =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
# take out private information
    $altered =~ s/\[\s*privado[^]]*\]\s*//gx;

    # take out private information
    $fh_string = $fh_string . $altered . $family{$key} . "\n";
    $worksheet_index->write_string( $row++, 0, $altered . $family{$key} );
}

open $fh, ">:encoding(UTF-8)", $dir . "mfs_index.txt" or croak $!;
print $fh $fh_string;

# print "String:\n\n$fh_string\n";
close $fh;

$workbook->close();

print "Done.\n";
