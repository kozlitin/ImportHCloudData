#!perl 

use strict;

use Spreadsheet::XLSX;
use Text::CSV;

my $excel = Spreadsheet::XLSX -> new($ARGV[0]);

my $csv = Text::CSV->new ( { binary => 1, eol => $/ } ) or die "Cannot use CSV: ".Text::CSV->error_diag ();

open my $fh, ">", "result.csv" or die "result.csv: $!";

foreach my $sheet (@{$excel -> {Worksheet}}) {

	next unless $sheet->{Name} =~ /^\d+\s*-/;
	
	my $Client = $sheet->{Cells}[2][1]->{Val};
	my $ResponsiblePerson = $sheet->{Cells}[5][1]->{Val};
	my $Currency = $sheet->{Cells}[0][40]->{Val};
	
	$csv->print($fh, [$sheet->{Name}, $Client, $ResponsiblePerson, $Currency]);
	
}

 close $fh or die "result.csv: $!";