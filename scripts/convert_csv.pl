#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use Data::Dumper;
use Text::CSV_XS;
use Term::ProgressBar;

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my @files = qw(
  AllMembers
  Camp
  Childcare
  Companies
  MembershipMapping
  MembershipOrders
  PrdRates
  ProgramCodes
  Programs
);

foreach my $filename (@files) {
  my($dataFile, $headers, $totalRows) = open_data_file('exports/' . $filename . '.csv');
  open(my $outFile, '>:encoding(UTF-8)', 'data/' . $filename . '.txt')
    or die "Couldn't open data/$filename.txt: $!";

  print $outFile join("\t", @{$headers}) . "\n";

  print "Converting $filename\n";
  my $progress = Term::ProgressBar->new({ 'count' => $totalRows });

  my $columnCount = scalar(@{$headers});
  my $count = 1;
  while(my $rowIn = $csv->getline($dataFile)) {

    $progress->update($count++);

    print $outFile join("\t", @{$rowIn}) . "\n";
  }

  close($dataFile);  
}