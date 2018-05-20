#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Date::Manip;
use Text::CSV;
use Term::ProgressBar;

my $templateName = 'DCT_ORDER_MBR_ASSOCIATE-41924';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ASSOCIATE_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'MemberId' },
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
  perMemberId
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

  $familyOrders->{$values->{'familyId'}} = {
    'orderNo' => $values->{'orderNo'},
    'memberId' => $values->{'shipCustomerId'},
    'billingId' => $values->{'billCustomerId'}
  }
}

close($orderMaster);

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

my($membersFile, $headers, $totalRows) = open_members_file();

print "Processing customers\n";
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });

my $count = 1;
my $row = 1;
while(my $rowIn = $csv->getline($membersFile)) {

  $progress->update($count++);

  my $values = clean_customer(map_values($headers, $rowIn));
  # dump($values); exit;

  my $familyId = $values->{'FamilyId'};

  next unless (exists($familyOrders->{$familyId}));
  next if ($values->{'MemberId'} eq $familyOrders->{$familyId}{'billingId'});
  
  $values->{'OrderNo'} = $familyOrders->{$familyId}{'orderNo'};

  write_record(
    $worksheet,
    $row++,
    make_record($values, \@allColumns, $columnMap)
  );
}
close($membersFile);
