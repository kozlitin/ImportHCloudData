#!perl 

use strict;

use Win32::OLE;
use Text::CSV;

my $Period = $ARGV[1];

if (!$Period) {
	print "Period value should be specified!";
	exit(1);
}

my $USDExchangeRate = $ARGV[2];

if (!$USDExchangeRate) {
	print "USD exchange rate should be specified!";
	exit(1);
}

my $ExcelOle = Win32::OLE->new('Excel.Application', 'Quit');
my $ExcelBookOle = $ExcelOle->Workbooks->Open($ARGV[0],,1);

if (!$ExcelBookOle) {
	print "Can not open workbook $ARGV[0]\n";
	$ExcelOle->Quit();
	$ExcelOle = undef;
	exit(1);
}

my $csv = Text::CSV->new ( { binary => 1, eol => $/ } ) or die "Cannot use CSV: ".Text::CSV->error_diag ();

open my $fh, ">", "result.csv" or die "result.csv: $!";

my $Counter = 0;

for (my $i=1; $i <= $ExcelBookOle->Sheets->{Count}; $i++ ) {

	my $sheet = $ExcelBookOle->Worksheets($i);

	next unless $sheet->{Visible};
	
	next unless $sheet->{Name} =~ /^\d+\s*-/;
	
	next unless ($sheet->Cells(7,40)->{Value} =~ /Âàðò³ñòü Ïîñëóãè íà ì³ñÿöü/);
		
	my $Client = $sheet->Cells(3,2)->{Value};
	my $ResponsiblePerson = $sheet->Cells(6,2)->{Value};
	my $Currency = $sheet->Cells(1,41)->{Value};
	
	my $Category = '';
	my $Item;
		
	foreach my $row (1 .. $sheet->Cells->SpecialCells(11)->{Row}) {
	 
		my $paragraph = $sheet->Cells($row,2)->{Value};
		
		if ($paragraph =~ /^5\.\d+/) {
			$Category = '-';
		} elsif ($paragraph =~ /^\d+\.\d+/) {
			$Category = $sheet->Cells($row,7)->{Value};
		}
		
		my ($price_text, $Price, $PriceUAH, $PriceUSD, $item_temp);
		
		if ($paragraph =~ /^\s*$/ || $paragraph =~ /^5\.\d+/) {
			$price_text = $sheet->Cells($row,40)->{Value};
			next if $price_text =~ /-/;
			next unless $price_text;
			$Price = $price_text+0;
			if ($Price != 0) {
				$item_temp = $sheet->Cells($row,3)->{Value};
				if ($item_temp) {
					$Item = $item_temp;
				}
				if ($Currency =~ /ÃÐÍ/) {
					$PriceUAH = $Price;
					$PriceUSD = sprintf("%.2f", $Price/$USDExchangeRate)+0;
				} else {
					$PriceUAH = sprintf("%.2f", $Price*$USDExchangeRate)+0;
					$PriceUSD = $Price;				
				}
				$csv->print($fh, [$Counter++, $Period, $sheet->{Name}, $Client, $ResponsiblePerson, $Currency, $Category, $Item, $Price, $PriceUAH, $PriceUSD]);
			}
		}
		
	}	
	
}

 close $fh or die "result.csv: $!";
 
$ExcelBookOle->Close();

$ExcelOle->Quit();
$ExcelOle = undef;
