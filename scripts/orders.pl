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
$csv->print($orderMaster, [member_order_master_fields()]);

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
      $values->{'PerMemberId'},
      $values->{'SponsorDiscount'},
    ]);
  }
) if (1);

close($orderMaster);

print Dumper($unmappedMembers) if (@{$unmappedMembers});

open(my $programMaster, '>', 'data/program_orders.csv')
  or die "Couldn't open data/program_orders.csv: $!";
$csv->print($programMaster, [program_order_master_fields()]);

my $orderHeaderMap = {
  'Session' => 'Session',
  'Program End Date' => 'ProgramEndDate',
  'Last Name' => 'LastName',
  'Item Description' => 'ItemDescription',
  'Member ID' => 'MemberId',
  'Receipt Number' => 'ReceiptNumber',
  'Fee Paid' => 'FeePaid',
  'Date Paid' => 'DatePaid',
  'GL Account' => 'GlAccount',
  'Program Start Date' => 'ProgramStartDate',
  'Billable Member Last Name' => 'BillableLastName',
  'Billable Member First Name' => 'BillableFirstName',
  'Branch' => 'Branch',
  'Branch Name' => 'BranchName',
  'Cycle' => 'Cycle',
  'Program Description' => 'ProgramDescription',
  'First Name' => 'FirstName',
  'billable Member Id' => 'BillableMemberId'  
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

    $csv->print($programMaster, [
      @{$record},
      $values->{'Session'},
      $values->{'ProgramEndDate'},
      $values->{'LastName'},
      $values->{'ItemDescription'},
      $values->{'MemberId'},
      $values->{'ReceiptNumber'},
      $values->{'FeePaid'},
      $values->{'DatePaid'},
      $values->{'GlAccount'},
      $values->{'ProgramStartDate'},
      $values->{'BillableLastName'},
      $values->{'BillableFirstName'},
      $values->{'Branch'},
      $values->{'BranchName'},
      $values->{'Cycle'},
      $values->{'ProgramDescription'},
      $values->{'FirstName'},
      $values->{'BillableMemberId'},
      $values->{'OrderNo'},
      $values->{'OrderDate'},
      $values->{'StatusDate'},
      $values->{'PerMemberId'},
      $values->{'PerBillableMemberId'},
      ]);
  },
  undef,
  $orderHeaderMap
);
