#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;

my $templateName = 'DCT_ORDER_MASTER-42939';

my $orderNo = 1000000000;

my $columnMap = {
  'ORDER_NO'                      => { 'type' => 'record', 'source' => 'orderNo' },
  'ORDER_DATE'                    => { 'type' => 'record', 'source' => 'joinDate' },
  'ORG_ID'                        => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                   => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'BILL_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'billableId' },
  'BILL_ADDRESS_TYPE_CODE'        => { 'type' => 'static', 'source' => 'HOME' },
  'SHIP_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'memberId' },
  'ORDER_STATUS_CODE'             => { 'type' => 'static', 'source' => 'A' },
  'ORDER_STATUS_DATE'             => { 'type' => 'record', 'source' => 'statusDate' },
  'APPLICATION'                   => { 'type' => 'static', 'source' => 'ORD001' },
  'FND_GIVE_EMPLOYER_CREDIT_FLAG' => { 'type' => 'static', 'source' => 'N' },
  'POS_FLAG'                      => { 'type' => 'static', 'source' => 'N' },
};

my @skipTypes = (
  'Family Program Participant',
  'HEALTH CTR LIFE',
  'LIFE MEMBER',
  'LIFE MEMBER FAM HC UPGRADE',
  'LIFE MEMBER HC FAM PLUS UPGRAD',
  'LIFE MEMBER/HEALTH CTR',
  'Non-Member',
  'PM FAMILY PLUS',
  'Program Membership Family',
  'Program Membership Individual',
  'Program Participant',
  'RETIREE - INDIVIDUAL',
  'RETIREE - UPGRADE FAMILY',
  'Retiree',
  'SAGE/PS PROGRAM INDIVIDUAL',
  'SAGE/PS PROGRAM UPGRADE FAMILY',
  'TRADE - FAMILY',
  'TRADE - INDIVIDUAL',
  'TRADE - TWO ADULT HH',
  'TRADE FAM UPGRADE FAM PLUS',
  'TRADE-INDIV UPGRADE INDIV PLUS',
  'Teen Summer Pass',
  'Trade HC Upgrade TWO Adult HC',
  'Trade HC Upgrade Two Adult HC',
);

# Read in billable family member data
my $billableFamilyMembers = {};
open(my $in, '<', 'data/billable members.txt')
  or die "Couldn't open data/billable members.txt: $!";
<$in>; # eat the headers
while (<$in>) {
  chomp;
  my($memberId, $familyId) = split("\t");
  $billableFamilyMembers->{$familyId} = $memberId;
}
close($in);

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

open(my $orderMaster, '>', 'data/order_master.txt')
  or die "Couldn't open data/order_master.txt: $!";
print $orderMaster join("\t", @allColumns, 'NEXT_BILL_DATE', 'JOIN_DATE') . "\n";;

my %types;
open(my $membersFile, '<', 'data/all members.txt')
  or die "Couldn't open data/all members.txt: $!";
<$membersFile>; # eat the headers

my $row = 1;
while(<$membersFile>) {
  chomp;
  my $values = split_values($_, qw(memberId familyId type nextBillDate joinDate statusDate));

  $types{$values->{'type'}}++;

  next if (grep { $_ eq $values->{'type'} } @skipTypes);

  unless (exists($billableFamilyMembers->{$values->{'familyId'}})) {
    print "Can't find billable ID for $values->{'memberId'}/$values->{'familyId'} "
      . "$values->{'type'}\n";
    next;
  }

  $values->{'billableId'} = $billableFamilyMembers->{$values->{'familyId'}};
  $values->{'orderNo'} = $orderNo++;

  my $record = make_record($values, \@allColumns, $columnMap);
  write_record($worksheet, $row++, $record);

  print $orderMaster join("\t", @{$record}, $values->{'nextBillDate'}, $values->{'joinDate'}) . "\n";;

}

close($membersFile);

foreach my $type (sort keys %types) {
  print "$type\t$types{$type}\n";
}
