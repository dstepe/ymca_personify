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

my $templateName = 'DCT_ORDER_MASTER-42939';

my $columnMap = {
  'ORDER_NO'                      => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_DATE'                    => { 'type' => 'record', 'source' => 'OrderDate' },
  'ORG_ID'                        => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                   => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'BILL_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'PerBillableMemberId' },
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

open(my $orderMaster, '>', 'data/member_orders.csv')
  or die "Couldn't open data/member_orders.csv: $!";
$csv->print($orderMaster, [member_order_fields()]);

my $types = {};
my $unmappedMembers = [];

my $orderNo = 1000000000;
my $order = 1;
process_data_file(
  'data/MembershipOrders.csv',
  sub {
    my $values = shift;
    # print Dumper($values);exit;

    $types->{$values->{'MembershipTypeDes'}}{$values->{'PaymentMethod'}}++;

    return if (grep { uc $_ eq uc $values->{'MembershipTypeDes'} } @skipTypes);

    $values->{'PerMemberId'} = lookup_id($values->{'MemberId'});
    $values->{'PerBillableMemberId'} = $values->{'PerMemberId'};

    unless ($values->{'PerMemberId'}) {
      push(@{$unmappedMembers}, $values);
      return;
    }

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

    $record = make_record($values, [member_order_fields()], make_column_map([member_order_fields()]));
    $csv->print($orderMaster, $record);
  }
) if (1);

close($orderMaster);

print Dumper($unmappedMembers) if (@{$unmappedMembers});

open(my $programMaster, '>', 'data/program_orders.csv')
  or die "Couldn't open data/program_orders.csv: $!";
$csv->print($programMaster, [program_order_fields()]);

my $orderHeaderMap = {
  'billable Member Id' => 'BillableMemberId',
  'Branch Name' => 'BranchName',
  'Cycle' => 'Cycle',
  'Date Paid' => 'DatePaid',
  'Fee Paid' => 'FeePaid',
  'Item Description' => 'ItemDescription',
  'Member ID' => 'MemberId',
  'Program Description' => 'ProgramDescription',
  'Program End Date' => 'ProgramEndDate',
  'Program Start Date' => 'ProgramStartDate',
  'Receipt Number' => 'ReceiptNumber',
  'Session' => 'Session',
};

process_data_file(
  'data/ProgramOrders.csv',
  sub {
    my $values = shift;
    
    $values->{'OrderNo'} = $orderNo++;

    $values->{'PerMemberId'} = lookup_id($values->{'MemberId'});
    $values->{'PerBillableMemberId'} = lookup_id($values->{'BillableMemberId'});

    $values->{'OrderDate'} = UnixDate($values->{'DatePaid'}, '%Y-%m-%d');
    $values->{'StatusDate'} = $values->{'OrderDate'};

    $values->{'FeePaid'} =~ s/\$//;

    my $record = make_record($values, \@allColumns, $columnMap);
    write_record($worksheet, $order++, $record);

    $record = make_record($values, [program_order_fields()], make_column_map([program_order_fields()]));
    $csv->print($programMaster, $record);
  },
  undef,
  $orderHeaderMap
) if (1);

$orderHeaderMap = {
  'Amount' => 'FeePaid',
  'Branch Name' => 'BranchName',
  'Date' => 'Date',
  'End Date' => 'ProgramEndDate',
  'Location' => 'Location',
  'Member Fee' => 'MemberFee',
  'Member ID' => 'MemberId',
  'Non- Member Fee' => 'NonMemberFee',
  'Participant Id' => 'ParticipantId',
  'Primary Sponsor Id' => 'PrimarySponsorId',
  'Program Branch' => 'ProgramBranch',
  'Program Description' => 'ProgramDescription',
  'Program Location Or Session' => 'ProgramLocation',
  'Program Member Fee' => 'ProgramMemberFee',
  'Program Schedule' => 'Schedule',
  'Program Type' => 'ProgramType',
  'Program' => 'Program',
  'Receipt Number' => 'ReceiptNumber',
  'Season' => 'Season',
  'Schedule' => 'Schedule',
  'Start Date' => 'ProgramStartDate',
};

# process_data_file(
#   'data/ChildcareOrders.csv',
#   sub {
#     my $values = shift;

#     $values->{'OrderNo'} = $orderNo++;

#     $values->{'PerMemberId'} = lookup_id($values->{'ParticipantId'});
#     $values->{'PerBillableMemberId'} = lookup_id($values->{'PrimarySponsorId'});

#     $values->{'OrderDate'} = UnixDate($values->{'Date'}, '%Y-%m-%d');
#     $values->{'StatusDate'} = $values->{'OrderDate'};

#     $values->{'FeePaid'} =~ s/\$//;

#     my $record = make_record($values, \@allColumns, make_column_map([program_order_fields()]));
#     write_record($worksheet, $order++, $record);

#     $csv->print($programMaster, $record);
#   },
#   undef,
#   $orderHeaderMap
# );

sub make_column_map {
  my $headers = shift;

  my $map = {};

  foreach my $key (@{$headers}) {
    $map->{$key} = {
      'type' => 'record',
      'source' => $key,
    }
  }

  return $map;
}