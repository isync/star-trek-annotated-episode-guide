#!perl


use Path::Tiny;
use Spreadsheet::XLSX;
use utf8; 

for my $file ("voyager.xlsx","the-next-generation.xlsx"){
	my $excel = Spreadsheet::XLSX->new($file);

	my @lines;
	foreach my $sheet (@{$excel->{Worksheet}}) {
		print "Sheet: ". $sheet->{Name} ."\n";
      
	       $sheet->{MaxRow} ||= $sheet->{MinRow};
	       foreach my $row (2 .. $sheet->{MaxRow}) {

			print "row:$row = ";
			$sheet->{MaxCol} ||= $sheet->{MinCol};

			next if  $sheet->{Cells}[$row][2] eq ""; ## skip season delimiter

			foreach my $col (2,3,5,11) {
				my $cell = $sheet->{Cells}[$row][$col];
				if ($cell) {
					print "$col:". $cell->{Val} .", ";
				}
			}
			print "\n";

			my $episode_no	= $sheet->{Cells}[$row][2]->{Val};
			my $episode	= $sheet->{Cells}[$row][3]->{Val};
			my $title	= $sheet->{Cells}[$row][5]->{Val};
			my $tags	= $sheet->{Cells}[$row][11]->{Val};

			$tags =~ s/^\s+|\s+$//g;
			$tags = lc($tags);

			my $line = "\t<tr>";
			$line   .= "<td>". $episode_no ."</td>";
			$line   .= "<td>". $episode ."</td>";
			$line   .= "<td>". $title ."</td>";
			$line   .= "<td>". $tags ."</td>";
			$line   .= "</tr>\n";

			push(@lines, $line);
	       }
	}

	for(@lines){
#		print $_;
	}

	my $outfile = "st-". $file;
	$outfile =~ s/\.xlsx/\.html/;

	my @contents = path($outfile)->lines_raw;
	my @contents2 = @contents;
	my ($begin,$end);
	for my $i (0 .. $#contents){
		$begin = $i if $contents[$i] =~ /-- TABLE_BEGIN --/;
		$end = $i if $contents[$i] =~ /-- TABLE_END --/;
	}
	print "$outfile, replace begin at $begin, end at $end \n";

	my @output = ( splice(@contents, 0, $begin +1), @lines, splice(@contents2, $end) );


	path($outfile)->spew_raw(@output); # writes to temp file then renames over
}
