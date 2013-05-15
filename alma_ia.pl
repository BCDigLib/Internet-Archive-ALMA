#!C:/Perl/bin/perl -w
use strict;

use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use IO::File;
use utf8;
use Cwd;


my $excelfile=shift(@ARGV);
my $dir = getcwd;
$dir=~s/\//\\/g;
print "dir is $dir\n";
$excelfile=$dir."//".$excelfile;
                                              

#Open file for output
my $outputfile = "url".$ARGV[0];

my $fh = IO::File->new($outputfile, 'w')
	or die "unable to open output file for writing: $!";
binmode($fh, ':utf8');

#Call MARC reformatter
my $records_hash=MARCreformatter();

#Call add links
add_links($records_hash);

#Write output
foreach my $k (sort keys %$records_hash) 
{
	$fh->print($$records_hash{$k}."\n");
}

#Close output file
$fh->close();

#########
#########
sub add_links 
########
{
	my $records_hash=shift;

	#Run Win32
	$Win32::OLE::Warn = 3; # Die on Errors.
	# ::Warn = 2; throws the errors, but #
	# expects that the programmer deals  #

	#First, we need an excel object to work with, so if there isn't an open one, we create a new one, and we define how the object is going to exit

	my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

	#For the sake of this program, we'll turn off all those pesky alert boxes, such as the SaveAs response "This file already exists", etc. using the 	#DisplayAlerts property.

	$Excel->{DisplayAlerts}=0;   
	my $Book = $Excel->Workbooks->Open($excelfile);   

	#Create a reference to a worksheet object and activate the sheet to give it focus so that actions taken on the workbook or application objects 	#occur on this sheet unless otherwise specified.

	my $Sheet = $Book->Worksheets(1);
	$Sheet->Activate();  

	#Find Used Range of Worksheet

	my $usedRange = $Sheet->UsedRange()->{Value};

	my $LastRow = $Sheet->UsedRange->Find({What=>"*",
    	SearchDirection=>xlPrevious,
    	SearchOrder=>xlByRows})->{Row};

	print "last row = $LastRow\n";

	my $LastCol = $Sheet->UsedRange->Find({What=>"*", 
                  SearchDirection=>xlPrevious,
                  SearchOrder=>xlByColumns})->{Column};
	print "last col is $LastCol\n";
   
	#Iterate through records hash and add links

	foreach my $k (sort keys %$records_hash) 
	{

		foreach my $row (@$usedRange) ##read excel rows
		{     			
			my ($wflink, $shipment, $physical_item, $digital_item, $identifier, $sysno, $barcode, $title, $vol, $author, $year, $callno, $notes) = @$row;
			if ($sysno && ($sysno ne "BIB ID")) {$sysno=$sysno}
			if ($sysno && ($sysno eq $k)) ## add 856
			{            
				if ($vol) 
				{
					$$records_hash{$k}="$$records_hash{$k}"."=856  40\$3$vol\$uhttp:\/\/www.archive.org\/details\/".$identifier."\n"
				}
				else 
				{
					$$records_hash{$k}="$$records_hash{$k}"."=856  40\$uhttp:\/\/www.archive.org\/details\/".$identifier."\n"
				}
			}
		}
	}

	return $records_hash;
};


#########
#########
sub MARCreformatter {
#########
	#calls subroutines to fix specific fields, and deletes, modifies, and adds some standard fields in the record
	my ($sec,$min,$hour,$mday,$mon,$yr,$wday,$yday,$isdst)=localtime();

	#change PERL default record delimiter
	$/="\n\n";

		
	my %records_hash = ();	

	while (<>) #here, ARGV is the MARC file
	{ 		
		my $oclcno;
		my $orig_sysno=0;
		my $call_no_date=0;
		#print "dollar under is $_\n";
		chomp;
		my $record=$_."\n";
		$record =~/=001\s\s(\d*)\n/;   ##get sys no from MARC record
		$orig_sysno=$1;
		print "system number from Marc record is $orig_sysno\n";
				
#find specific fields, change them here or in subroutine
			my @record_parts = split(/\n/, $record);
			foreach my $record_part (@record_parts) {
				if ($record_part =~ m/=LDR/){ #change status to new, cat rules to isbd punc, and enc lvl to K
					$record_part =~ m/(.{11}).(.{11})..(.*)/;
					$record_part=$1."n".$2."Ki".$3;
				}

				if ($record_part =~ m/245\s\s\d\d\$a/) {	#identify title
					print "main sub title is $record_part\n";
					$record_part=title($record_part);
				} 
				if ($record_part =~ m/=008  .{7}(\d{4})/){ #find date in 008 to add to call no and form o
					$call_no_date=$1;	
					$record_part =~ m/(.{6}).{6}(.{17}).(.*)/;
					$record_part=$1.sprintf("%02d",$yr-100).sprintf("%02d",$mon+1).sprintf("%02d",$mday).$2."o".$3;
				}
				if ($record_part =~ m/=090|=050/){	#add date to call #
					$record_part =~ s/=050.../=090   /;
					if ($record_part !~ m/\d\d\d\d$/) {$record_part=$record_part." ".$call_no_date."eb";} else {$record_part=$record_part."eb"}
				}
				if ($record_part =~ m/=035  \\\\\$a\(OCoLC\)/){	#identify OCLC # for 776
					$oclcno=$';
				}
				if ($record_part =~ m/=260/) {	#identify pub details and call sub routine
					#print "260 is $record_part\n";
					$record_part=pubdetails($record_part);
				} 
				if ($record_part =~ m/=300/) {	#identify phys desc and call sub routine
					#print "300 is $record_part\n";
					$record_part=physdesc($record_part);
				} 



        		}

#Eliminate some fields from the record
			@record_parts = grep {!/^=951|^=950|^=952|^=948|^=004|^=010|^=005|^=040|^=042|^=035|^=966|^=099|^=049|^=092|^=949|^=999|^=019|^=994|^=901|^=001|^=029|^=590|^=533|^=539/} @record_parts;
#iterate through record array and write to file

			$record = join("\n", @record_parts);

		

#Add some fields to the record
			$record=$record."\n";
			$record=$record."=006  m        d        \n";
			$record=$record."=007  cr uuu---uuuuu\n";
			$record=$record."=040  \\\\\$aBXM\$cBXM\n";
			$record=$record."=533  \\\\\$aElectronic reproduction.\$bSan Francisco, Calif. :\$cOpen Content Alliance,\$d".(1900+$yr).".\$nMode of access: Internet.\n";
			$record=$record."=776  18\$cOriginal\$w(Alma)".$orig_sysno."\$w\(OCoLC\)".$oclcno."\n";
			$record=$record."=940  \\\\\$aOCA\n";
		


			$records_hash{ $orig_sysno } = $record; 

			}

		return (\%records_hash);
};

#########
sub title {
#########
#title sub adds |h and isbd punctuation
#########
	my $title=shift;

	print "title is $title\n";

# add subfield c delimiter when there is a slash spaces on either side
	if ($title =~ m/\s\/\s/) {
		$title=~s/ \/ / \/\$c/;
		print "WARNING -- sub c added\n";
		}
# fix old fashioned alternative title -- need to add 246 manually
	if ($title =~ m/;\$bor/) {
		$title=~s/;\$bor/, or/;
		print "WARNING -- ALT TITLE Corrected\n";
		}
# lower case certain words after |c
	if ($title =~ m/cEdit|cWith |cBy /) {
		$title =~s/(\$c.)/lc($1)/e;
	print "LOOOOOOK\n";
	}
# no |b or |c
	if ($title !~ m/\$b|\$c/) {
		if ($title =~ m/\w\.\w\.$/)
			{$title=$title.'$h[electronic resource].'}
		else {$title=~s/\W*$/\$h[electronic resource]./};
		print "title is $title\n";
		print "NOT B OR C\n\n";
	}
# |c, no |b

	if ($title =~ m/\$c/ && $title!~m/\$b/){
		$title=~s/[,\s\.\/]*\$c/\$h[electronic resource] \/\$c/;
		print "title is $title\n";
		print "C, but NOT B \n\n";
	}

# |b
	if ($title =~ m/\$b/){
		$title=~s/[;:,\s\.]+\$b/\$h[electronic resource] :\$b/;
		$title=~s/!\$b/!\$h[electronic resource] :\$b/;
		$title=~s/\?\$b/\?\$h[electronic resource] :\$b/;
		$title=~s/[;,\s\.\/]*\$c/ \/\$c/;
		print "title is $title\n";
		print "B \n\n";
	}
		
	return($title);

}

#########
sub pubdetails {
#########

	my $pubdetails=shift;
	print "$pubdetails\n";
	if ($pubdetails =~m/[:,]\$a/) {$pubdetails =~s/[:,]\$a/ ;\$a/}
	if ($pubdetails =~m/[a-z]\$a/) {$pubdetails =~s/([a-z])\$a/$1 ;\$a/}
	if ($pubdetails =~m/\]\$b/) {$pubdetails =~s/\$b/ :\$b/g}
	if ($pubdetails =~m/\$b/) {$pubdetails =~s/,\$b/ :\$b/g}
	if ($pubdetails =~m/(\w|\])\$c/) {$pubdetails =~s/\$c/,\$c/}

	print "$pubdetails\n\n";
	return($pubdetails);



}
#########
sub physdesc {
#########

	my $physdesc=shift;
	print "$physdesc\n";
	if (($physdesc =~m/\$b/) && ($physdesc!~m/:\$b/)) {$physdesc =~s/\$b/ :\$b/}
	if (($physdesc =~m/\$c/) && ($physdesc!~m/;\$c/)) {$physdesc =~s/\$c/ ;\$c/}
	print "$physdesc\n\n";
	return($physdesc);



}
=pod

use: ia.pl picklist.xls records.mrk

Takes an Internet Archive EXCEL pick list and adds urls to aleph records in .mrk format. 
The following line of the script must be adjusted if picklist column names vary: my ($shipment , $physical_item, $identifier, $sysno, $barcode, $vol, $year, $title, $author, $callno, $notes) = @$row;
Handling rejected items on picklists is out of scope of the script.
 
Outputs a file with 'url' prefixed to the name of the original .mrk file.

mckelvee@bc.edu October 1, 2009

Updates + Notes 20091227 
1.) revised to work with Wonderfetch format; deals with empty rows in wonderfetch
2.) now looks for first worksheet in excel workbook, rather than a sheet named "Sheet1"
3.)detects current working directory and expects to find excel sheet there
4.) change call number suffix to eb
5.) The script should be handling unicode diacritics, but it isn't so we are outputting MARC 8 and running the script on that.  The problem may be that our opac using combining rather than combined unicode diacritics.  

Updated December 31, 2011 -- deals with duplicate records. 

Didn't
2.) 300 |a 1 online resource with page numbers in parens -- http://www.loc.gov/catdir/pcc/sca/FinalVendorGuide.pdf


=cut