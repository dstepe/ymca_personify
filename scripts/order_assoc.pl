#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Date::Manip;
use Text::CSV;

my $templateName = 'DCT_ORDER_MBR_ASSOCIATE-41924';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'orderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ASSOCIATE_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'memberId' },
  'ASSOCIATE_CLASS_CODE'               => { 'type' => 'static', 'source' => 'FAMILY' },
};

my @orderMasterFields = qw(
  orderNo
  orderDate
  orgId
  orgUnitId
  billCustomerId
  billAddressTypeCode
  shipCustomerId
  orderMethodCode
  orderStatusCode
  orderStatusDate
  ordstsReasonCode
  clOrderMethodCode
  couponCode
  application
  ackLetterMethodCode
  poNumber
  confirmationNo
  ackLetterPrintDate
  confirmationDate
  orderCompleteFlag
  advContractId
  advAgencyCustId
  billSalesTerritory
  fndGiveEmployerCreditFlag
  shipSalesTerritory
  posFlag
  posCountryCode
  posState
  posPostalCode
  advRateCardYearCode
  advAgencySubCustId
  employerCustomerId
  oldOrderNo
  membershipType
  paymentMethod 
  renewalFee
  branchCode
  membershipBranch
  companyName
  nextBillDate
  joinDate
  familyId
);

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

open(my $orderMaster, '<', 'data/order_master.txt')
  or die "Couldn't open data/order_master.txt: $!";
<$orderMaster>; # eat the headers

my $familyOrders = {};
while(<$orderMaster>) {
  chomp;
  my $values = split_values($_, @orderMasterFields);
  #print Dumper($values); exit;

  $familyOrders->{$values->{'familyId'}} = {
    'orderNo' => $values->{'orderNo'},
    'memberId' => $values->{'shipCustomerId'},
  }
}

close($orderMaster);

$/ = "\r\n";

my $csv = Text::CSV->new();

open(my $members, '<:encoding(UTF-8)', 'data/AllMembers.csv')
  or die "Couldn't open data/AllMembers.csv: $!";
<$members>;

my $row = 1;
while (my $line = <$members>) {
  chomp $line;

  $csv->parse($line) || die "Line could not be parsed: $line";

  my($memberId, $familyId) = $csv->fields();

  next unless (exists($familyOrders->{$familyId}));

  my $values = {
    'orderNo' => $familyOrders->{$familyId}{'orderNo'},
    'memberId' => $memberId,
  };

  write_record(
    $worksheet,
    $row++,
    make_record($values, \@allColumns, $columnMap)
  );
}
close($members);
