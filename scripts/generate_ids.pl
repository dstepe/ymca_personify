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

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my @idAllColumns = qw( TRX_ID PERSONIFY_ID TYPE );

my $idWorkbook = make_workbook('id_map');
my $idWorksheet = make_worksheet($idWorkbook, \@idAllColumns);

my $customerSeq = 500000;
my $companySeq = 100000;

my $row = 1;

$dbh->do(q{
  delete from ids
  });

$dbh->do(q{
  delete from companies
  });

my %companies;

process_data_file(
  'data/Companies.csv',
  sub {
    my $values = shift;

    my $id = sprintf('%012d', $companySeq++);

    write_record($idWorksheet, $row++, [$values->{'TRX_ID'}, $id, 'company']);
    $dbh->do(q{
      insert into ids (t_id, p_id)
        values (?, ?)
      }, undef, $values->{'TRX_ID'}, $id);

    $companies{$values->{'TRX_ID'}} = $id;
  }
);

process_data_file(
  'data/AllMembers.csv',
  sub {
    my $values = shift;
    
    my $id = sprintf('%012d', $customerSeq++);

    return if (exists($companies{$values->{'MemberId'}}));

    write_record($idWorksheet, $row++, [$values->{'MemberId'}, $id, 'person']);
    $dbh->do(q{
      insert into ids (t_id, p_id)
        values (?, ?)
      }, undef, $values->{'MemberId'}, $id);
  }
);
