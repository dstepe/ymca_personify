#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::CSV_XS;
use Term::ProgressBar;

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my @idAllColumns = qw( TRX_ID PERSONIFY_ID TYPE );

my $idWorkbook = make_workbook('id_map');
my $idWorksheet = make_worksheet($idWorkbook, \@idAllColumns);

open(my $idMap, '>', 'data/id_map.txt')
  or die "Couldn't open data/id_map.txt: $!";
$csv->print($idMap, ['TrxId', 'PersonifyId', 'Type']);

my $customerSeq = 500000;
my $companySeq = 100000;

my($dataFile, $headers, $totalRows) = open_data_file('data/AllMembers.csv');

print "Processing customers\n";
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });

my $count = 1;
my $row = 1;
while(my $rowIn = $csv->getline($dataFile)) {

  $progress->update($count++);

  my $values = map_values($headers, $rowIn);
  
  write_record($idWorksheet, $row++, [$values->{'MemberId'}, $customerSeq++, 'person']);
  $csv->print($idMap, [$values->{'MemberId'}, $customerSeq, 'person']);

}

close($dataFile);

($dataFile, $headers, $totalRows) = open_data_file('data/Companies.csv');

print "Processing companies\n";
$progress = Term::ProgressBar->new({ 'count' => $totalRows });

$count = 1;
while(my $rowIn = $csv->getline($dataFile)) {

  $progress->update($count++);

  my $values = map_values($headers, $rowIn);

  write_record($idWorksheet, $row++, [$values->{'CompanyId'}, $companySeq++, 'company']);
  $csv->print($idMap, [$values->{'CompanyId'}, $companySeq, 'company']);

}

close($dataFile);

