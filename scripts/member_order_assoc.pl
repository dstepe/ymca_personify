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

my($ordersFile, $headers, $totalRows) = open_data_file('data/member_orders.csv');

my $familyOrders = {};
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });
my $row = 1;
my $count = 1;
while(my $rowIn = $csv->getline($ordersFile)) {

  $progress->update($count++);

  my $values = map_values($headers, $rowIn);

  $familyOrders->{$values->{'FamilyId'}} = {
    'OrderNo' => $values->{'OrderNo'},
    'MemberId' => $values->{'ShipCustomerId'},
    'BillingId' => $values->{'BillCustomerId'}
  }
}
close($ordersFile);

$row = 1;
process_customer_file(
  sub {
    my $values = shift;

    my $familyId = $values->{'FamilyId'};

    return unless (exists($familyOrders->{$familyId}));
    return if ($values->{'PerMemberId'} eq $familyOrders->{$familyId}{'BillingId'});
    
    $values->{'OrderNo'} = $familyOrders->{$familyId}{'OrderNo'};

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);
