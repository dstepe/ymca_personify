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

our $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my @idAllColumns = qw( TRX_ID PERSONIFY_ID TYPE );

my $idWorkbook = make_workbook('id_map');
my $idWorksheet = make_worksheet($idWorkbook, \@idAllColumns);

my $customerSeq = 500000;
my $companySeq = 100000;

my $row = 1;

my %companies;

process_data_file(
  'data/Companies.csv',
  sub {
    my $values = shift;

    my $id;
    if (has_p_id($values->{'TRX_ID'})) {
      ($id) = $dbh->selectrow_array(q{
        select p_id
          from ids
          where t_id = ?
        }, undef, $values->{'TRX_ID'});
    } else {
      $id = sprintf('%012d', $companySeq++);

      $dbh->do(q{
        insert into ids (t_id, p_id)
          values (?, ?)
        }, undef, $values->{'TRX_ID'}, $id);
    }

    write_record($idWorksheet, $row++, [$values->{'TRX_ID'}, $id, 'company']);
    $companies{$values->{'TRX_ID'}} = $id;
  }
);

process_data_file(
  'data/AllMembers.csv',
  sub {
    my $values = shift;
    
    return if (exists($companies{$values->{'MemberId'}}));

    my $id;
    if (has_p_id($values->{'MemberId'})) {
      ($id) = $dbh->selectrow_array(q{
        select p_id
          from ids
          where t_id = ?
        }, undef, $values->{'MemberId'});
    } else {
      $id = sprintf('%012d', $customerSeq++);

      $dbh->do(q{
        insert into ids (t_id, p_id)
          values (?, ?)
        }, undef, $values->{'MemberId'}, $id);
    }

    write_record($idWorksheet, $row++, [$values->{'MemberId'}, $id, 'person']);
  }
);

sub has_p_id {
  my $t_id = shift;

  our $dbh;

  my($idExists) = $dbh->selectrow_array(q{
    select count(*)
      from ids
      where t_id = ?
    }, undef, $t_id);

  return $idExists;
}