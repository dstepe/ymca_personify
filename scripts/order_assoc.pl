#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Date::Manip;
use Text::CSV_XS;
use Term::ProgressBar;
use DBI;

my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

my $templateName = 'DCT_ORDER_MBR_ASSOCIATE-41924';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ASSOCIATE_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'PerMemberId' },
  'ASSOCIATE_CLASS_CODE'               => { 'type' => 'static', 'source' => 'FAMILY' },
};

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

# Load assoc orders file
my $assocOrder = {};
my($assocOrdersFile, $headers, $totalRows) = open_data_file('data/assoc_orders.csv');
while(my $rowIn = $csv->getline($assocOrdersFile)) {
  my $values = map_values($headers, $rowIn);

  my $membershipType = uc $values->{'MembershipType'};

  unless (exists($assocOrder->{$values->{'BillingMemberId'}}{$values->{'FamilyId'}}{$membershipType})) {
    $assocOrder->{$values->{'BillingMemberId'}}{$values->{'FamilyId'}}{$membershipType} = [];
  }

  push(@{$assocOrder->{$values->{'BillingMemberId'}}{$values->{'FamilyId'}}{$membershipType}}, $values->{'MemberId'});

}
close($assocOrdersFile);

my $memberOrderProblemsWorkbook = make_workbook('memberOrder_problems');
my $memberOrderProblemsWorksheet = make_worksheet($memberOrderProblemsWorkbook, 
  ['OrderTrxMemberId', 'OrderPerMemberId', 'FamilyId', 'CustomerTrxFamilyBillable', 
    'CustomerPerFamilyBillable', 'Problem']);
my $problemRow = 1;

# For each member order, find any assocs and add them here
my $ordersFile;
($ordersFile, $headers, $totalRows) = open_data_file('data/member_orders.csv');

my $familyOrders = {};
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });
my $row = 1;
my $count = 1;
while(my $rowIn = $csv->getline($ordersFile)) {

  $progress->update($count++);

  my $values = map_values($headers, $rowIn);

  my $membershipType = uc $values->{'MembershipTypeDes'};
  next if ($values->{'AccessDenied'} eq 'Deny');

  if (!exists($assocOrder->{$values->{'PerBillableMemberId'}})) {
     write_record($memberOrderProblemsWorksheet, $problemRow++, [
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      $values->{'FamilyId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'Order billable ID is not expected to be a billable ID',
    ]);
  } elsif (!exists($assocOrder->{$values->{'PerBillableMemberId'}}{$values->{'FamilyId'}})) {
     write_record($memberOrderProblemsWorksheet, $problemRow++, [
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      $values->{'FamilyId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'Order billable ID does not match family billable ID',
    ]);
  } elsif (!exists($assocOrder->{$values->{'PerBillableMemberId'}}{$values->{'FamilyId'}}{$membershipType})) {
     write_record($memberOrderProblemsWorksheet, $problemRow++, [
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      $values->{'FamilyId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'Order membership type ' . $values->{'MembershipTypeDes'} . ' is not found in family membership types',
    ]);
  }

  my $members = $dbh->selectcol_arrayref(q{
    select p_id
      from members
      where f_id = ?
        and membership = ?
    }, undef, $values->{'FamilyId'}, $membershipType);

  foreach my $assocMember (@{$members}) {
    # Primary members are part of the main order and not added here
    next if ($assocMember eq $values->{'PerBillableMemberId'});

    my $record = {
      'OrderNo' => $values->{'OrderNo'},
      'PerMemberId' => $assocMember,
    };

    write_record(
      $worksheet,
      $row++,
      make_record($record, \@allColumns, $columnMap)
    );
  }

}
close($ordersFile);
