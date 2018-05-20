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

my $templateName = 'DCT_ORDER_MASTER-42939';

my $orderNo = 1000000000;

my $columnMap = {
  'ORDER_NO'                      => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_DATE'                    => { 'type' => 'record', 'source' => 'OrderDate' },
  'ORG_ID'                        => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                   => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'BILL_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'PerMemberId' },
  'BILL_ADDRESS_TYPE_CODE'        => { 'type' => 'static', 'source' => 'HOME' },
  'SHIP_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'PerMemberId' },
  'ORDER_STATUS_CODE'             => { 'type' => 'static', 'source' => 'A' },
  'ORDER_STATUS_DATE'             => { 'type' => 'record', 'source' => 'StatusDate' },
  'APPLICATION'                   => { 'type' => 'static', 'source' => 'ORD001' },
  'FND_GIVE_EMPLOYER_CREDIT_FLAG' => { 'type' => 'static', 'source' => 'N' },
  'POS_FLAG'                      => { 'type' => 'static', 'source' => 'N' },
};

my @skipTypes = (
  'COLLEGE SUMMER',
  'DIABETES 6 MTH FAM UPGRADE',
  'DIABETES 6 MTH INDIVIDUAL',
  'EMPLOY UPGRADE FAM',
  'EMPLOY UPGRADE FAM PLUS',
  'EMPLOY UPGRADE FAMILY',
  'EMPLOY UPGRADE HC',
  'EMPLOY UPGRADE HC FAM',
  'EMPLOY UPGRADE IND PLUS',
  'EMPLOY UPGRADE INDIV PLUS',
  'EMPLOY UPGRADE INDIV PLUS DEP',
  'EMPLOY UPGRADE SP FAM',
  'EMPLOY UPGRADE TWO ADULT',
  'EMPLOY YOUNG ADULT',
  'EMPLOYEE UPGRADE HC FAM PLUS',
  'FAMILY PROGRAM PARTICIPANT',
  'FAMILY PROGRAM PARTICIPANT',
  'HEALTH CTR LIFE',
  'LIFE MEMBER',
  'LIFE MEMBER FAM HC UPGRADE',
  'LIFE MEMBER HC FAM PLUS UPGRAD',
  'LIFE MEMBER/HEALTH CTR',
  'PM FAMILY PLUS',
  'PROGRAM MEMBERSHIP FAMILY',
  'PROGRAM MEMBERSHIP INDIVIDUAL',
  'RETIREE - INDIVIDUAL',
  'RETIREE - UPGRADE FAMILY',
  'RETIREE',
  'SAGE/PS PROGRAM INDIVIDUAL',
  'SAGE/PS PROGRAM UPGRADE FAMILY',
  'TEEN SUMMER PASS',
  'XXXCINCINNATI UPGRADE',
  'XXXSENIOR ADULT - EMPLOYEE',
);

my $cycleDurations = {
  'Annual' => '1 year',
  'Monthly E-Pay' => '1 month',
  'Quarterly' => '3 months',
};

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

open(my $orderMaster, '>', 'data/order_master.txt')
  or die "Couldn't open data/order_master.txt: $!";
$csv->print($orderMaster, [order_master_fields()]);

my $types = {};

my($ordersFile, $headers, $totalRows) = open_data_file('data/MembershipOrders.csv');

print "Processing orders\n";
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });
my $order = 1;
my $count = 1;
while(my $rowIn = $csv->getline($ordersFile)) {

  $progress->update($count++);

  my $values = map_values($headers, $rowIn);
  # print Dumper($values);exit;

  $types->{$values->{'MembershipTypeDes'}}{$values->{'PaymentMethod'}}++;

  next if (grep { uc $_ eq uc $values->{'MembershipTypeDes'} } @skipTypes);

  $values->{'PerMemberId'} = convert_id($values->{'MemberId'});

  # OrderDate must be start of current membership cycle
  #  NextBillDate - (method offset)
  $values->{'OrderDate'} = UnixDate(
    DateCalc(
      ParseDate($values->{'NextBillDate'}), 
      '-' . $cycleDurations->{$values->{'PaymentMethod'}}
      ), 
    '%Y-%m-%d'
  );

  # Catch any future dates

  $values->{'StatusDate'} = $values->{'OrderDate'};
  $values->{'OrderNo'} = $orderNo++;

  my $record = make_record($values, \@allColumns, $columnMap);
  write_record($worksheet, $order++, $record);

  $values->{'RenewMembershipFee'} =~ s/[^0-9\.]//g;

  $csv->print($orderMaster, [
    @{$record}, 
    $values->{'MembershipTypeDes'},
    $values->{'PaymentMethod'}, 
    $values->{'RenewMembershipFee'}, 
    $values->{'BranchCode'}, 
    $values->{'MembershipBranch'}, 
    $values->{'CompanyName'},
    $values->{'NextBillDate'}, 
    $values->{'JoinDate'},
    $values->{'FamilyId'}, 
    $values->{'PerMemberId'}
    ]);

}

close($ordersFile);
close($orderMaster);

# foreach my $type (sort keys %{$types}) {
#   foreach my $method (sort keys %{$types->{$type}}) {
#     print "$type\t$method\t$types->{$type}{$method}\n";
#   }
# }
