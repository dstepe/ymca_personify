#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::CSV_XS;
use Term::ProgressBar;
use DBI;

my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

$dbh->do(q{
  delete from companies
  });

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

my($dataFile, $headers, $totalRows) = open_data_file('data/CompanyData.csv');

our $programCodes = {};
while(my $rowIn = $csv->getline($dataFile)) {

  my $values = map_values($headers, $rowIn);

  $dbh->do(q{
    insert into companies (t_id, p_id, c_name)
      values (?, ?, ?)
    }, undef, $values->{'TRX_ID'}, $values->{'CUSTOMER_ID'}, $values->{'NAME'});

}

close($dataFile);
