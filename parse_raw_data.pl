#!perl 

use strict;

use Win32::OLE;
use Text::CSV;
use IO::File;
use XML::Writer;
use Getopt::Long;
use File::Basename;
use File::Spec;
use Cwd;
use Encode qw(encode decode is_utf8);

my $ExcelFileName;

GetOptions (  "excelfilename=s"  => \$ExcelFileName
			  ) or &{ 
					print_usage();
					exit(1)
					};

if (!$ExcelFileName) {
	print "Excel file name should be specified!\n";
	print_usage();
	exit(1);
}

if(dirname($ExcelFileName) =~ /^\.$/) {
	$ExcelFileName = File::Spec->catfile( getcwd , $ExcelFileName );
}

my $ExcelOle = Win32::OLE->new('Excel.Application', 'Quit');
my $ExcelBookOle = $ExcelOle->Workbooks->Open($ExcelFileName);

if (!$ExcelBookOle) {
	print "Can not open workbook $ARGV[0]\n";
	$ExcelOle->Quit();
	$ExcelOle = undef;
	exit(1);
}

my $csv = Text::CSV->new ( { binary => 1, eol => $/ } ) or die "Cannot use CSV: ".Text::CSV->error_diag ();

open my $fh, ">", "result.csv" or die "result.csv: $!";

my $writer = XML::Writer->new( OUTPUT => 'self', DATA_MODE => 1 );
$writer->xmlDecl("UTF-8");
$writer->startTag("records");

###
# Getting EDRPOU table from first sheet
#
my %EDRPOUTable;

my $sheet = $ExcelBookOle->Worksheets(1);

my $EDRPOUColumn = 8;

if ($sheet->Cells(5,7)->{Value} =~ /ªÄÐÏÎÓ/) {
	$EDRPOUColumn = 7;
}

foreach my $row (1 .. $sheet->Cells->SpecialCells(11)->{Row}) {
	next unless $sheet->Cells($row,6)->{Value} =~ /^1$/;
	$EDRPOUTable{$sheet->Cells($row,3)->{Value}} = $sheet->Cells($row,$EDRPOUColumn)->{Value};
}

###
#
#

my $Counter = 0;
my $MonthlyPriceColumn = 0;

for (my $i=1; $i <= $ExcelBookOle->Sheets->{Count}; $i++ ) {

	my $sheet = $ExcelBookOle->Worksheets($i);

	next unless $sheet->{Visible};
	
	next unless $sheet->{Name} =~ /^\d+\s*-/;

	if ($sheet->Cells(7,40)->{Value} =~ /Âàðò³ñòü Ïîñëóãè íà ì³ñÿöü/) {
		$MonthlyPriceColumn = 40;
	} elsif ($sheet->Cells(7,22)->{Value} =~ /Âàðò³ñòü Ïîñëóãè íà ì³ñÿöü/) {
		$MonthlyPriceColumn = 22;
	} elsif ($sheet->Cells(7,42)->{Value} =~ /Âàðò³ñòü Ïîñëóãè íà ì³ñÿöü/) {
		$MonthlyPriceColumn = 42;	
	} else {
		next;
	}
			
	my $Client = $sheet->Cells(3,2)->{Value};
	my $ResponsiblePerson = $sheet->Cells(6,2)->{Value};
	
	my ($currency_temp, $Currency);
	
	my $currency_temp = $sheet->Cells(1,$MonthlyPriceColumn+1)->{Value};
	if ($currency_temp =~ /ÃÐÍ/) {
		$Currency = "980";
	} elsif ($currency_temp =~ /ÄÎËÀÐ ÑØÀ/) {
		$Currency = "840";
	} 

	unless($Currency) {
		if ($sheet->{Name} =~ /\$/) {
			$Currency = "840";
		}		
	}
	
	unless($Currency) {
		$Currency = "980";
	}
	
	print $sheet->{Name} . "\n";	
		
	my %SheetData;	
		
	ExtractDataFromSheet($sheet, \%SheetData, 1);
	
	for my $discount_offset (0 .. 9) {
		$sheet->Cells( 1, 9 + $discount_offset*3 )->{Value} = 0;
	}
	
	$sheet->Calculate();
	
	ExtractDataFromSheet($sheet, \%SheetData, 0);

	for my $row_num (keys %SheetData) {	
		
		my $edrpou = $EDRPOUTable{$Client};
		
		next unless ($edrpou);
		
		$csv->print($fh, [$Counter++, $sheet->{Name}, $Client, $edrpou, $ResponsiblePerson, $Currency, $SheetData{$row_num}->{Category}, $SheetData{$row_num}->{Service}, $SheetData{$row_num}->{Qty}, $SheetData{$row_num}->{Unit}, ($SheetData{$row_num}->{Price} ? $SheetData{$row_num}->{Price} : 0), ($SheetData{$row_num}->{Price0} ? $SheetData{$row_num}->{Price0} : 0)]);
		
		$writer->startTag("record");
		$writer->dataElement( counter => $Counter );
		$writer->dataElement( sheetname => decode('windows-1251', $sheet->{Name}) );
		$writer->dataElement( client => decode('windows-1251', $Client) );
		$writer->dataElement( edrpou => $edrpou );
		$writer->dataElement( responsibleperson => decode('windows-1251', $ResponsiblePerson) );
		$writer->dataElement( currency => $Currency );
		$writer->dataElement( category => decode('windows-1251', $SheetData{$row_num}->{Category}) );
		$writer->dataElement( service => decode('windows-1251', $SheetData{$row_num}->{Service}) );
		$writer->dataElement( qty => $SheetData{$row_num}->{Qty} );
		$writer->dataElement( unit => decode('windows-1251', $SheetData{$row_num}->{Unit}) );
		$writer->dataElement( price => ($SheetData{$row_num}->{Price} ? $SheetData{$row_num}->{Price} : 0) ) ;
		$writer->dataElement( price0 => ($SheetData{$row_num}->{Price0} ? $SheetData{$row_num}->{Price0} : 0) );
		$writer->endTag("record");		
		
	}	
		
}

close $fh or die "result.csv: $!";
 
$writer->endTag("records");

open my $fhxml, ">:encoding(UTF-8)", "result.xml" or die "result.xml: $!";
print $fhxml $writer->to_string;
close $fhxml or die "result.xml: $!";
 
$ExcelBookOle->Close(0);

$ExcelOle->Quit();
$ExcelOle = undef;

######################################################
# Subroutines
######################################################

sub print_usage {
	print "Usage: perl $0 --excelfilename=excel_file_name\n";
}

sub ExtractDataFromSheet {

	my ($sheet, $DataHashRef, $mode) = @_;

	my $Category = '';
	my $Item;	
	
	foreach my $row (1 .. $sheet->Cells->SpecialCells(11)->{Row}) {
	 
		my $paragraph = $sheet->Cells($row,2)->{Value};
		
		if ($paragraph =~ /^5\.\d+/) {
			$Category = '-';
		} elsif ($paragraph =~ /^\d+\.\d+/) {
			$Category = $sheet->Cells($row,7)->{Value};
		}
		
		my ($price_text, $price_onetime_text, $Price, $PriceUAH, $PriceUSD, $item_temp, $Unit, $Qty, $qty_temp);
		
		if ($paragraph =~ /^\s*$/ || $paragraph =~ /^5\.\d+/ || $paragraph =~ /^\-$/) {
			$price_text = $sheet->Cells($row,$MonthlyPriceColumn)->{Value};
			$price_onetime_text = $sheet->Cells($row,$MonthlyPriceColumn-3)->{Value};
			$Price = ($price_text+0)+($price_onetime_text+0);
			next unless $Price;
			if ($Price != 0) {
				$item_temp = $sheet->Cells($row,3)->{Value};
				if ($item_temp) {
					$Item = $item_temp;
				}
				$qty_temp = $sheet->Cells($row,7)->{Value};
				if ($qty_temp =~ /^\d+$/) {
					$Qty = $qty_temp;
				} else {
					$Qty = 1;
				}
				$Unit = $sheet->Cells($row,4)->{Value};
				$DataHashRef->{$row}->{Category} = $Category;
				$DataHashRef->{$row}->{Service} = $Item;
				$DataHashRef->{$row}->{Qty} = $Qty;
				$DataHashRef->{$row}->{Unit} = $Unit;
				if ($mode) {
					$DataHashRef->{$row}->{Price} = $Price;
				} else {
					$DataHashRef->{$row}->{Price0} = $Price;
				}	
			}
		}
		
	}	
	
	
}