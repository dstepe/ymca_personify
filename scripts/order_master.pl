#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Date::Manip;
use Text::CSV;

my $templateName = 'DCT_ORDER_MASTER-42939';

my $orderNo = 1000000000;

my $columnMap = {
  'ORDER_NO'                      => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_DATE'                    => { 'type' => 'record', 'source' => 'OrderDate' },
  'ORG_ID'                        => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                   => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'BILL_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'MemberId' },
  'BILL_ADDRESS_TYPE_CODE'        => { 'type' => 'static', 'source' => 'HOME' },
  'SHIP_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'MemberId' },
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

open(my $orderMaster, '>', 'data/order_master.txt')
  or die "Couldn't open data/order_master.txt: $!";
print $orderMaster join("\t", @allColumns, 'MEMBERSHIP_TYPE', 
  'PAYMENT_METHOD', 'RENEWAL_FEE', 'BRANCH_CODE', 'BRANCH_NAME',
  'COMPANY_NAME', 'NEXT_BILL_DATE', 'JOIN_DATE', 'FAMILY_ID') . "\n";;

my $csv = Text::CSV->new();

my $types = {};
# Read in membership order data
$/ = "\r\n";

open(my $orders, '<:encoding(UTF-8)', 'data/MemberShipOrders.csv')
  or die "Couldn't open data/MemberShipOrders.csv: $!";
my $headerLine = <$orders>;
$csv->parse($headerLine) || die "Line could not be parsed: $headerLine";
my @headers = $csv->fields();

my $order = 1;
while(my $line = <$orders>) {
  chomp $line;

  $csv->parse($line) || die "Line could not be parsed: $line";

  my $values = map_values(\@headers, [$csv->fields()]);
  #print Dumper($values);exit;

  $types->{$values->{'MembershipTypeDes'}}{$values->{'PaymentMethod'}}++;

  next if (grep { uc $_ eq uc $values->{'MembershipTypeDes'} } @skipTypes);

  # OrderDate must be start of current membership cycle
  #  NextBillDate - (method offset)
  $values->{'OrderDate'} = UnixDate(
    DateCalc(
      ParseDate($values->{'NextBillDate'}), 
      '-' . $cycleDurations->{$values->{'PaymentMethod'}}
      ), 
    '%Y-%m-%d'
  );

  # Catch an future dates

  $values->{'StatusDate'} = $values->{'OrderDate'};
  $values->{'OrderNo'} = $orderNo++;

  my $record = make_record($values, \@allColumns, $columnMap);
  write_record($worksheet, $order++, $record);

  $values->{'RenewMembershipFee'} =~ s/[^0-9\.]//g;

  print $orderMaster join("\t", @{$record}, $values->{'MembershipTypeDes'},
    $values->{'PaymentMethod'}, $values->{'RenewMembershipFee'}, 
    $values->{'BranchCode'}, $values->{'MembershipBranch'}, 
    $values->{'CompanyName'},
    $values->{'NextBillDate'}, $values->{'JoinDate'},
    $values->{'FamilyId'}) . "\n";

}

close($orders);

# foreach my $type (sort keys %{$types}) {
#   foreach my $method (sort keys %{$types->{$type}}) {
#     print "$type\t$method\t$types->{$type}{$method}\n";
#   }
# }
