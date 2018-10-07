# mfs-directory
From google-sheet Roster to .xlsx through perl to LibreOffice document and pdf booklet 

The .odt is the editable word-processing file that was
used, mostly pulling directly via the Roster and perl programs. So,
make many corrections on Roster, then one effort to convert the current
Roster into a spreadsheet for perl and then onto the Directory. Here
is my first draft of the directions to accomplish that effort:

Correct the google sheet "MFSRoster2018-2019".
File... Download as... Microsoft Excel (.xlsx) to your "Downloads" folder.
In a command window:
  ls -la ~/Downloads/MFSRoster2018-2019.xlsx
  perl mfs-directory.pl
  ls -la ~/Downloads/store
  perl volunteers-mfs.pl
  ls -la directory.xls

Edit in Libre Office a file such as Directorio_2018-2019-v1.odt
  Tools...Update... Update all
  Open mfs_index.txt and copy the entire contents, then use it to
replace the Indices
  Insert... Page Break (CTRL-return) to correct widows
  Tools...Update... Update all (to correct Table of Contents)
  Check that last page is evenly divisible by 4, so that back page of
Booklet is correct.  If needed, correct the length of the Notes
Section.
  Tools...Update... Update all (to correct Table of Contents)
  Save
  File... Export as PDF... Export

Run the PDF-Booklet software.
  Files... Files... Downloads... The Name ... Open Selected Files...
 Notice the left page number is a multiple of 4.
 Go

Email! Print! Put on web page! Lather, rinse, and repeat.
