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

my $companySeq = 100000;
my $customerSeq = 500000;

my $row = 1;

my %companies;
my %newIds;

my($maxCompanyId) = $dbh->selectrow_array(q{
  select max(p_id)
    from ids
    where p_id > ?
      and p_id < ?
  }, undef, format_p_id($companySeq), format_p_id($customerSeq));

my($maxCustomerId) = $dbh->selectrow_array(q{
  select max(p_id)
    from ids
    where p_id > ?
  }, undef, format_p_id($customerSeq));

$companySeq = $maxCompanyId + 1 if ($maxCompanyId > $companySeq);
$customerSeq = $maxCustomerId + 1 if ($maxCustomerId > $customerSeq);

print "Starting company id is $companySeq\n";
print "Starting customer id is $customerSeq\n";

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
      $id = format_p_id($companySeq++);

      $dbh->do(q{
        insert into ids (t_id, p_id)
          values (?, ?)
        }, undef, $values->{'TRX_ID'}, $id);

      $newIds{'company'}++;
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
      $id = format_p_id($customerSeq++);

      $dbh->do(q{
        insert into ids (t_id, p_id)
          values (?, ?)
        }, undef, $values->{'MemberId'}, $id);
      
      $newIds{'customer'}++;
    }

    write_record($idWorksheet, $row++, [$values->{'MemberId'}, $id, 'person']);
  }
);

print Dumper(\%newIds);

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

sub format_p_id {
  my $pId = shift;

  return sprintf('%012d', $pId);
}