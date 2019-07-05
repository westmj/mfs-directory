#!/usr/bin/perl -w
# for commmand line interface, maybe need to take out the -wT options, tainting hard to avoid...
# on my HP chromebook, try    $  cp ~/mfs-directory/cli-cgi.pl ~/public_html/cgi-bin/cli-cgi.pl 
use strict;
use warnings;
my $version = "2019-07-05-16:00";
use Carp 'verbose';
# local $SIG{__DIE__} = sub { Carp::confess(@_) };
use Data::Dumper;
use CGI::Simple;
use CGI::Carp 'fatalsToBrowser';
use Storable;
use Unicode::Collate;
use Spreadsheet::Read;
use Spreadsheet::ParseXLSX;
use Spreadsheet::WriteExcel;
use Getopt::Long;
use Archive::Zip qw( :ERROR_CODES :CONSTANTS ) ;
#use Sereal::Encoder qw(encode_sereal 
#sereal_encode_with_object
#);
#use Sereal::Decoder qw(decode_sereal 
#sereal_decode_with_object scalar_looks_like_sereal
#);
 
# sudo cpanm  Sereal::Encoder Sereal::Decoder Data::Dumper CGI::Simple CGI::Carp Storable Unicode::Collate Spreadsheet::Read Spreadsheet::ParseXLSX Spreadsheet::WriteExcel Getopt::Long Archive::Zip 

#    use Config;    # OS dependent configuration section

my ( $dir, $q , $book_copy );
if    ( $^O =~ /^MSWin32/xi ) { $dir = '/home/westmj/public_html/files/'; }
elsif ( $^O =~ /^linux/xi )   { $dir = '/home/westmj/public_html/files/'; }
elsif ( $^O =~ /^darwin/xi )  { $dir = '/Users/Headofschool/Downloads/'; }
my $book_store    = 'mfs_store';
my $workbook_name = 'directory.xls';

if ( defined( $ENV{'REQUEST_METHOD'} ) ) {
    if ( $ENV{'REQUEST_METHOD'} eq 'GET' ) {
        get();
    }
    elsif ( $ENV{'REQUEST_METHOD'} eq 'POST' ) {
        post();
    }
    else { print "Where the hell am I?\n"; }
}
else {
    cli();
}

sub cli {
    GetOptions( 
# $ perl -wT cli-cgi.pl --dir=/home/westmj/public_html/files/ --book_store=store --workbook=directory.xls 
        "dir=s"        => \$dir,
        "book_store=s" => \$book_store,
        "workbook=s"   => \$workbook_name
    ) or die("Error in command line arguments (see  https://github.com/westmj/mfs-directory ) \n");
    print "In the cli subroutine now..( https://github.com/westmj/mfs-directory ) \$dir ='$dir' \$book_store = '$book_store' \$workbook_name = '$workbook_name'  .\n";
    my $book =
      ReadData( "$dir" . 'MFSRoster 2019-2020.xlsx' );   # xlsx;
    print "past... ReadData \n";
    my $taint_store = $dir . $book_store;
    my $untaint_store;
    $taint_store =~ /^([A-Za-z0-9_\/]+)$/;
    $untaint_store = $1; 
    # $untaint_store = $dir . $untaint_store;
    print "store = '$untaint_store'\n";

#my $encoder = Sereal::Encoder->new();
#my $out = $encoder->encode(\$book);
#$encoder->encode_to_file($untaint_store ,\$book );
 
     store \$book, $untaint_store;
#    $book_copy = retrieve( $dir . 'store' );
    print " book stored... "; 
    make_booklet_support(); 
    print " booklet made ... "; 
    exit;
}

sub post {
    print "Content-type: text/html\n\n";
    print '<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Post MFS Directory maker</title>
  <meta name="description" content="mwjwest@gmail.com">
  <meta name="author" content="westmj">
  <link rel="stylesheet" href="css/styles.css?v=1.0">
</head>
<body>
';
    $CGI::Simple::POST_MAX = 1024;        # max upload via post default 100kB
    $CGI::Simple::POST_MAX = 100000000;   # max upload via post default ~ 100 mB
    $CGI::Simple::DISABLE_UPLOADS = 0;    # enable uploads
    my ( $ok, );
    $q = CGI::Simple->new;
    print "<p>See also https://github.com/westmj/mfs-directory </p>";
    my %params = $q->Vars;                # as a plain hash
    print "More...<br> "
      . join( "  <br>   \n",
        Dumper(%params), ' ', $q->param('submit'), " ", " end " );

    if ( $q->param('submit') eq 'upload' ) {
        print "<br> Inside the upload section. <br>\n";
        $ok = $q->upload( $q->param('upload_file1'),
            $dir . $q->param('upload_file1') );
        if ($ok) {
            print "\n<br>Uploaded "
              . $q->param('upload_file1')
              . " and wrote it OK! ";
            print "<br> Inside the ok section. <br>with \n" . "$dir"
              . $q->param('upload_file1') . '<br>';

            #            if ($ok) {
            my $book = ReadData( "$dir" . $q->param('upload_file1') );    # xlsx
            print "<br> Past ReadData. <br>\n";
            store \$book, $dir . $book_store;
            print "<br> Past store. <br>\n";
              make_booklet_support(); 
              print "<br> Past booklet support. <br>\n";
         #$book_copy = retrieve( $dir . 'store' );
         #
         #                print "More..." . join( "  <br>   \n", Dumper(%ENV) );
         #            }
        }
        else {
            print "File " . $q->param('upload_file1') . " upload failed: \n",
              $q->cgi_error();
        }
        exit;
    }
    elsif ( $q->param('submit') eq 'copy' ) {
        print "<br> COPY! <br> \n";
    }
    else { print "submitted, but not an upload or copy request, sorry. \n"; }
    return ();
}

sub make_store {
    my $book =
      ReadData( "$dir" . $q->param('upload_file1') );    # MFSRoster2018-2019
    store \$book, $dir . $book_store;
    $book_copy = retrieve( $dir .  $book_store );

}

sub get {
    print "Content-type: text/html\n\n";
    print '<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>MFS Directory maker GET</title>
  <meta name="description" content="mwjwest@gmail.com">
  <meta name="author" content="westmj">
  <link rel="stylesheet" href="css/styles.css?v=1.0">
</head>
<body>
<b>MFS Directory maker from mwjwest@gmail.com</b>
<p>See also https://github.com/westmj/mfs-directory </p>
<p>
Select from your local computer storage, the ".xls" file (downloaded the Roster in Google Sheets), to process into a new Directory of families.<br>
</p>
<FORM
 METHOD="POST"
 ACTION="http://127.0.0.1/~westmj/cgi-bin/cli-cgi.pl"
 ENCTYPE="multipart/form-data">
    <INPUT TYPE="file" NAME="upload_file1" SIZE="142">
    <button type="submit" name="submit" value="upload">Submit</button>
</FORM>
<p>
or 
</p>
<FORM
 METHOD="POST"
 ACTION="http://127.0.0.1/~westmj/cgi-bin/cli-cgi.pl"
 ENCTYPE="multipart/form-data">
    <button type="submit" name="submit" value="copy">Get a new copy from the last data submitted.</button>
</FORM>';
print "<p> Version = '$version' </p>"; 
    print '
</body>
</html>';
    return ();
}

sub make_booklet_support {
    my (
        $fh,     $fh_string,   $sheet,       $father,
        $mother, $father_name, $mother_name, %family,
        @staff,  @committee,   @volunteers,  $table,
        $row,    $guardians,   $guardians_name, %parents_guardians, 
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

    #$decoder->decode_from_file($file);
#   my $decoder = Sereal::Decoder->new();
#$decoder->decode_from_file($file);
#my $structure;
#$decoder->decode($blob, $structure); # deserializes into $structure

    my $book_copy       = retrieve( $dir . $book_store );
    print "<p> past retrieve </p>\n"; 
    # Name the file that holds the excel workbook of worksheets for the directory
    my $workbook        = Spreadsheet::WriteExcel->new( $dir . $workbook_name );
   print "<p> past workbook </p>\n"; 
     my $worksheet       = $workbook->add_worksheet('staff');
        print "<p> past worksheet  </p>\n"; 
 
    my $worksheet_index = $workbook->add_worksheet('index');

    #staff
    foreach my $table ( 1 .. 2 ) {
        $sheet = $$book_copy->[$table]{cell};
        foreach my $row ( 2 .. scalar( @{$sheet} ) ) {
            if ( !$$book_copy->[$table]{cell}[1][$row] ) { next }
            if (   !$$book_copy->[$table]{cell}[1][$row]
                or !$$book_copy->[$table]{cell}[2][$row]
                or !$$book_copy->[$table]{cell}[16][$row]
                or !$$book_copy->[$table]{cell}[7][$row]
                or !$$book_copy->[$table]{cell}[4][$row] )
            {
                print
"Something is empty: In table '$table' grade $grade[$table] row '$row'  
               first name '$$book_copy->[$table]{cell}[1][$row]' 
               last name '$$book_copy->[$table]{cell}[2][$row]' 
               phone '$$book_copy->[$table]{cell}[7][$row]'
               email '$$book_copy->[$table]{cell}[4][$row]' \n";
            }
            push @staff,
              $$book_copy->[$table]{cell}[1][$row] . ' ' .     # first name
              $$book_copy->[$table]{cell}[2][$row] . ' (' .    # last name
              $$book_copy->[$table]{cell}[16][$row] . ')  '
              .    # annotation / position
              $$book_copy->[$table]{cell}[7][$row] . ' ' .    # phone
              $$book_copy->[$table]{cell}[4][$row] . "\n";    # email
        }
    }
     print "<p> before open '$dir' . mfs_staff.txt </p>\n"; 
    open $fh, ">:encoding(UTF-8)", $dir . "mfs_staff.txt" or croak $!;
     print "<p> past open </p>\n"; 
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
              $$book_copy->[$table]{cell}[1][$row] . ' ' .     # first name
              $$book_copy->[$table]{cell}[2][$row] . ' (' .    # last name
              $$book_copy->[$table]{cell}[16][$row] . ')  '
              .    # annotation / position
              $$book_copy->[$table]{cell}[7][$row] . ' ' .    # phone
              $$book_copy->[$table]{cell}[4][$row] . "\n";    # email
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
              $$book_copy->[$table]{cell}[1][$row] . ' ' .     # first name
              $$book_copy->[$table]{cell}[2][$row] . ' (' .    # last name
              $$book_copy->[$table]{cell}[16][$row] . ')  '
              .    # annotation / position
              $$book_copy->[$table]{cell}[7][$row] . ' ' .    # phone
              $$book_copy->[$table]{cell}[4][$row] . "\n";    # email   ### empty problem  committee email ###
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
        $worksheet =
          $workbook->add_worksheet( $grade[$table] );    # name the table
        $sheet     = $$book_copy->[$table]{cell};        # get the sheet
        $fh_string = $fh_string . "$grade[$table]\n";

        foreach my $row ( 2 .. scalar( @{$sheet} ) ) {
            if ( !${ $$book_copy->[$table]{cell}[2] }[$row] ) {
                next;
            }                                            # empty line
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
                $father_name =~ s/(.*?),.*/$1/x; # chop off phone number and all

                unless ( exists( $family{$father} ) ) { $family{$father} = '' }
                $family{$father} =
                    $family{$father} . "\n"
                  . '                      '
                  . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
                  . ${ $$book_copy->[$table]{cell}[2] }[$row] . ' ('
                  . "$grade[$table])";           # add a kid
                $family{$father} =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;
                unless (exists ( $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } ) ) {
                    $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } = '' ;
                }
                 $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } =  $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } . "\t" . $father  ;
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
                unless (exists ( $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } ) ) {
                    $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } = '' ;
                }
                 $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } =  $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } . "\t" . $mother  ;
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
                
            
                unless (exists ( $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . 
                "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } ) ) {
                    $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . 
                    "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } = '' ;
                }
                 $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . 
                 "\t"  . ${ $$book_copy->[$table]{cell}[2] }[$row] } =  
                 $parents_guardians{ ${ $$book_copy->[$table]{cell}[1] }[$row] . "\t"  
                 . ${ $$book_copy->[$table]{cell}[2] }[$row] } . "\t" . $guardian   ;
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
                $father_name =~ s/(.*?),.*/$1/x; # chop off phone number and all

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
                $father_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;

                $mother_name = $mother;
                unless ($mother_name) { print "No mother name at row $row in table $table '$grade[$table]' ! \n";}
# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
                $mother_name =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x; ### empty problem mother name  ###
                $mother_name =~ s/(.*?),.*/$1/x; ### empty problem  mother name ###
                $fh_string =
                    $fh_string
                  . ${ $$book_copy->[$table]{cell}[1] }[$row] . ' '
                  . ${ $$book_copy->[$table]{cell}[2] }[$row]  ### empty problem ###
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

    open $fh, ">:encoding(UTF-16)", $dir . "mfs_index-utf16.txt" or croak $!;
    print $fh $fh_string;

    close $fh;

# parents_guardians
    $fh_string = '';
    $row       = 0;
    foreach my $key ( sort keys(%parents_guardians) ) {
        my $altered = $key;

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
        $altered =~ s/.([^\(\s]*?)\ \((.)(\1)\)/$2$3/x;

# look for those pesky names that start with an accented letter and are hard to alphabetize like Angel (Ángel)
# take out private information
        $altered =~ s/\[\s*privado[^]]*\]\s*//gx;

        # take out private information
        $fh_string = $fh_string .  $parents_guardians{$key} . "\n";
        $worksheet_index->write_string( $row++, 0, $altered . $parents_guardians{$key} );
    }

    open $fh, ">:encoding(UTF-16)", $dir . "mfs_parents_guardians-utf16.txt" or croak $!;
    print $fh $fh_string;

    close $fh;


    $workbook->close();
#    return(); 
    # create a zip archive to allow download of the new result 
    my $zip = Archive::Zip->new();
    # Add a file from disk
    my $file_member ;
     $file_member = $zip->addFile( $dir.'directory.xls', 'directory.xls' );
     $file_member = $zip->addFile( $dir.'mfs_index-utf16.txt', 'mfs_index-utf16.txt' ); 
     $file_member = $zip->addFile( $dir.'mfs_out.txt', 'mfs_out.txt' ); 
     $file_member = $zip->addFile( $dir.'mfs_committee.txt', 'mfs_committee.txt' ); 
     $file_member = $zip->addFile( $dir.'mfs_volunteers.txt', 'mfs_volunteers.txt' ); 
     $file_member = $zip->addFile( $dir.'mfs_staff.txt', 'mfs_staff.txt' ); 
     $file_member = $zip->addFile( $dir.'mfs_store', 'mfs_store' ); 
     print "<p> past mfs_store </p> \n"; 
     $file_member = $zip->addFile( $dir.'mfs_Directorio_2018-2019-v4.odt', 'mfs_Directorio_2018-2019-v4.odt'  ); 
      print "<p> past mfs_Directorio_2018-2019-v4.odt </p> \n"; 
 
# Save the Zip file
my $stamp = getLoggingTime(); 
unless ( $zip->writeToFileNamed($dir.'mfs_directory-'.$stamp.'.zip') == AZ_OK ) {
   die 'write error on zip file: ','mfs_directory-'.$stamp.'.zip';
}

print "Zip file: try something like <a href='/~westmj/files/".'mfs_directory-'.$stamp.'.zip'."'>mfs_directory-".$stamp.".zip</a> ";

    print "Done.\n";
}

sub getLoggingTime {
    my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst)=localtime(time);
    my $nice_timestamp = sprintf ( "%04d%02d%02d_%02d-%02d-%02d",
                                   $year+1900,$mon+1,$mday,$hour,$min,$sec);
    return $nice_timestamp;
}
