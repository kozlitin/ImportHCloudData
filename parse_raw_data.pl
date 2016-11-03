#!perl 

use strict;

use Spreadsheet::XLSX;
use Text::CSV;

my $excel = Spreadsheet::XLSX -> new($ARGV[0]);

my $Period = $ARGV[1];

if (!$Period) {
	my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
	my @abbr = qw(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec);
	$Period = sprintf("%s%d", $abbr[$mon], $year+1900);
}

my $csv = Text::CSV->new ( { binary => 1, eol => $/ } ) or die "Cannot use CSV: ".Text::CSV->error_diag ();

open my $fh, ">", "result.csv" or die "result.csv: $!";

my $Counter = 0;

foreach my $sheet (@{$excel -> {Worksheet}}) {

	next unless $sheet->{Name} =~ /^\d+\s*-/;
	
	my $Client = $sheet->{Cells}[2][1]->{Val};
	my $ResponsiblePerson = $sheet->{Cells}[5][1]->{Val};
	my $Currency = $sheet->{Cells}[0][40]->{Val};
	
	$sheet -> {MaxRow} ||= $sheet -> {MinRow};

	my $Category = '';
	my $Item;
	
	foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
	 
		my $paragraph = $sheet->{Cells}[$row][1]->{Val};
		
		if ($paragraph =~ /^5\.\d+/) {
			$Category = '-';
		} elsif ($paragraph =~ /^\d+\.\d+/) {
			$Category = $sheet->{Cells}[$row][6]->{Val};
		}
		
		my ($price_text, $Price, $item_temp);
		
		if ($paragraph =~ /^\s*$/ || $paragraph =~ /^5\.\d+/) {
			$price_text = $sheet->{Cells}[$row][39]->{Val};
			next if $price_text =~ /-/;
			next unless $price_text;
			$Price = $price_text+0;
			if ($Price != 0) {
				$item_temp = $sheet->{Cells}[$row][2]->{Val};
				if ($item_temp) {
					$Item = $item_temp;
				}
				$csv->print($fh, [$Counter++, $Period, $sheet->{Name}, $Client, $ResponsiblePerson, $Currency, $Category, $Item, $Price]);
			}
		}
		
	}	
	
}

 close $fh or die "result.csv: $!";