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
use MIME::Base64;
use Encode;
use DBI;

my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

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

open(my $allOrders, '>', 'data/all_orders.csv')
  or die "Couldn't open data/all_orders.csv: $!";
$csv->print($allOrders, [all_order_fields()]);

open(my $campOrdersBilling, '>', 'data/camp_orders_billing.csv')
  or die "Couldn't open data/camp_orders_billing.csv: $!";
$csv->print($campOrdersBilling, [all_order_fields()]);

open(my $orderMaster, '>', 'data/member_orders.csv')
  or die "Couldn't open data/member_orders.csv: $!";
$csv->print($orderMaster, [member_order_fields()]);

$dbh->do(q{
  update access set order_created = ?
  }, undef, 0);

my $types = {};
my $unmappedMembers = [];

my $orderNo = 1000000000;
my $order = 1;
process_data_file(
  'data/MembershipOrders.csv',
  sub {
    my $values = shift;

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

    $values->{'RenewMembershipFee'} =~ s/[^0-9\.]//g;
    $values->{'OrderTotal'} = $values->{'RenewMembershipFee'};
    $values->{'BalanceDue'} = 0;

    ($values->{'AccessDenied'}) = $dbh->selectrow_array(q{
      select access
        from access
        where p_id = ?
      }, undef, $values->{'PerMemberId'});

    if ($values->{'AccessDenied'} eq 'Deny') {
      $dbh->do(q{
        update access set order_created = ?
          where p_id = ?
        }, undef, 1, $values->{'PerMemberId'});
    }

    my $record = make_record($values, \@allColumns, $columnMap);
    write_record($worksheet, $order++, $record);

    $record = make_record($values, [member_order_fields()], make_column_map([member_order_fields()]));
    $csv->print($orderMaster, $record);
    $csv->print($allOrders, [
      $values->{'OrderNo'}, 
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'membership'
    ]);

  }
);

# Generate orders for access denied that do not have current orders
my $sth = $dbh->prepare(q{
  select t_id, p_id
    from access
    where order_created <> 1
      and access = 'Deny'
  });

$sth->execute();

my $currentDate = UnixDate(ParseDate('today'), '%Y-%m-%d');

while (my($denyForTid, $denyForPid) = $sth->fetchrow_array()) {
  my $values = {
    'AccessDenied' => 'Deny',
    'BranchCode' => '',
    'CompanyName' => '',
    'CurrentMembershipFee' => 0,
    'FamilyId' => '',
    'JoinDate' => $currentDate,
    'MemberId' => $denyForTid,
    'MembershipBranch' => 'Metropolitan',
    'MembershipTotal' => 0,
    'MembershipTypeDes' => 'Access Denied',
    'NextBillDate' => '12/31/2099',
    'OrderDate' => $currentDate,
    'OrderNo' => $orderNo++,
    'PaymentMethod' => 'Annual',
    'PerBillableMemberId' => $denyForPid,
    'PerMemberId' => $denyForPid,
    'RenewMembershipFee' => 0,
    'SponsorDiscount' => 0,
    'StatusDate' => $currentDate,
    'OrderTotal' => 0,
    'BalanceDue' => 0,
  };

  my $record = make_record($values, \@allColumns, $columnMap);
  write_record($worksheet, $order++, $record);

  $record = make_record($values, [member_order_fields()], make_column_map([member_order_fields()]));
  $csv->print($orderMaster, $record);
  $csv->print($allOrders, [
    $values->{'OrderNo'}, 
    lookup_t_id($values->{'PerMemberId'}),
    $values->{'PerMemberId'},
    lookup_t_id($values->{'PerBillableMemberId'}),
    $values->{'PerBillableMemberId'},
    'so_membership'
  ]);


  $dbh->do(q{
    update access set order_created = ?
      where p_id = ?
    }, undef, 1, $denyForTid);

}

close($orderMaster);

print Dumper($unmappedMembers) if (@{$unmappedMembers});

my $noProductCodeWorkbook = make_workbook('missing_product_code');
my $noProductCodeWorksheet = make_worksheet($noProductCodeWorkbook, 
  ['Type', 'Branch', 'Cycle', 'Description', 'Summary', 'Session', 'Program Skipped']);
my $noProductRow = 1;

open(my $programMaster, '>', 'data/program_orders.csv')
  or die "Couldn't open data/program_orders.csv: $!";
$csv->print($programMaster, [program_order_fields()]);

my $orderHeaderMap = {
  'billable Member Id' => 'BillableMemberId',
  'Branch Name' => 'BranchName',
  'Cycle' => 'Cycle',
  'Date Paid' => 'DatePaid',
  'Fee Paid' => 'OrderTotal',
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

    $values->{'BalanceDue'} = 0;

    $values->{'OrderTotal'} =~ s/[\$,]//g;

    $values->{'ProductCode'} = lookup_product_code('program', $values);

    unless ($values->{'ProductCode'}) {
      my $skipped = skip_program($values->{'ProgramDescription'});
      write_record($noProductCodeWorksheet, $noProductRow++, [
        'program',
        $values->{'BranchName'},
        $values->{'Cycle'},
        $values->{'ProgramDescription'},
        $values->{'ItemDescription'},
        '',
        $skipped ? 'Skipped' : ''
      ]);

      return;    
    }

    my $record = make_record($values, \@allColumns, $columnMap);
    write_record($worksheet, $order++, $record);

    $record = make_record($values, [program_order_fields()], make_column_map([program_order_fields()]));
    $csv->print($programMaster, $record);
    $csv->print($allOrders, [
      $values->{'OrderNo'}, 
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'program'
    ]);

  },
  undef,
  $orderHeaderMap
);

$orderHeaderMap = {
  'Branch' => 'BranchName',
  'Class Summary' => 'ItemDescription',
  'Current Balance' => 'BalanceDue',
  'Date Enrolled' => 'DateEnrolled',
  'Participant Id' => 'ParticipantId',
  'Primary Sponsor Id' => 'PrimarySponsorId',
  'Program Description' => 'ProgramDescription',
  'Program Fee' => 'OrderTotal',
  'Session' => 'Session',
  'Start Date' => 'ProgramStartDate',
};

process_data_file(
  'data/CampOrders.csv',
  sub {
    my $values = shift;

    $values->{'OrderNo'} = $orderNo++;

    $values->{'Cycle'} = '';
    $values->{'ProgramEndDate'} = '';
    $values->{'ReceiptNumber'} = '';
    
    $values->{'ProgramEndDate'} = UnixDate(
      DateCalc(ParseDate($values->{'ProgramStartDate'}), '+ 5 days'), 
      '%Y-%m-%d'
    );
    
    $values->{'DatePaid'} = $values->{'DateEnrolled'};

    $values->{'PerMemberId'} = lookup_id($values->{'ParticipantId'});
    $values->{'PerBillableMemberId'} = lookup_id($values->{'PrimarySponsorId'});

    $values->{'OrderDate'} = UnixDate($values->{'DateEnrolled'}, '%Y-%m-%d');
    $values->{'StatusDate'} = $values->{'OrderDate'};

    $values->{'OrderTotal'} =~ s/[\$,]//g;
    $values->{'BalanceDue'} =~ s/[\$,]//g;

    $values->{'ProductCode'} = lookup_product_code('camp', $values);

    unless ($values->{'ProductCode'}) {
      write_record($noProductCodeWorksheet, $noProductRow++, [
        'camp',
        $values->{'BranchName'},
        '',
        $values->{'ProgramDescription'},
        $values->{'ItemDescription'},
        $values->{'Session'}
      ]);

      return;     
    }

    my $record = make_record($values, \@allColumns, $columnMap);
    write_record($worksheet, $order++, $record);

    $record = make_record($values, [program_order_fields()], make_column_map([program_order_fields()]));
    $csv->print($programMaster, $record);
    $csv->print($allOrders, [
      $values->{'OrderNo'}, 
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'camp'
    ]);
    $csv->print($campOrdersBilling, [
      $values->{'OrderNo'}, 
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'}
    ]);

  },
  undef,
  $orderHeaderMap
);

close($programMaster);

$orderHeaderMap = {
  'memberid' => 'MemberId',
  'Attributions' => 'DonorName',
  'branch Name' => 'BranchName',
  'Billing Frequency' => 'BillingFrequency',
  'Campaign Balance' => 'CampaignBalance',
  'Campaign Description' => 'ItemDescription',
  'Campaign Pledge' => 'CampaignPledge',
  'Campaign Pledge Status' => 'CampaignPledgeStatus',
  'Corporation Name' => 'CorporationName',
  'Fair Market Value' => 'FairMarketValue',
  'Pledge Date' => 'PledgeDate',
  'Pledge ID' => 'PledgeID',
  'Pledge Next Bill Date' => 'PledgeNextBillDate',
  'Pledge Type' => 'PledgeType',
  'Pledge Type Frequency' => 'PledgeTypeFrequency',
  'Receipt Number' => 'ReceiptNumber',
  'Tracking Number' => 'TrackingNumber',
  'Volunteer Goal' => 'VolunteerGoal',
  'Volunteer Name' => 'VolunteerName',
};

open(my $donationMaster, '>', 'data/donation_orders.csv')
  or die "Couldn't open data/donation_orders.csv: $!";
$csv->print($donationMaster, [donation_order_fields()]);

process_data_file(
  'data/CampaignPledges.csv',
  sub {
    my $values = shift;

    $values->{'OrderNo'} = $orderNo++;

    $values->{'BranchName'} = resolve_branch_name($values)
      unless ($values->{'BranchName'});

    $values->{'Session'} = '';
    $values->{'ProgramStartDate'} = '';
    $values->{'ProgramEndDate'} = '';
    $values->{'OrderTotal'} = $values->{'CampaignPledge'};
    $values->{'BalanceDue'} = $values->{'CampaignBalance'};
    $values->{'DatePaid'} = UnixDate($values->{'PledgeDate'}, '%Y-%m-%d');
    $values->{'Cycle'} = '';

    $values->{'PerMemberId'} = lookup_id($values->{'MemberId'});
    $values->{'PerBillableMemberId'} = $values->{'PerMemberId'};

    $values->{'OrderDate'} = $values->{'PledgeDate'};
    $values->{'StatusDate'} = $values->{'PledgeDate'};

    $values->{'OrderTotal'} =~ s/[\$,]//g;
    $values->{'BalanceDue'} =~ s/[\$,]//g;

    $values->{'Comments'} =~ s/^\s+//;
    $values->{'Comments'} =~ s/\s+$//;
    $values->{'Comments'} = encode_base64($values->{'Comments'}, '');

    my $campaignProductDetails = lookup_campaign_product($values);

    $values->{'BranchName'} = $campaignProductDetails->{'BranchName'};
    $values->{'BranchCode'} = $campaignProductDetails->{'BranchCode'};
    $values->{'ProductCode'} = $campaignProductDetails->{'ProductCode'};
    $values->{'CampaignCode'} = $campaignProductDetails->{'CampaignCode'};
    $values->{'FundCode'} = $campaignProductDetails->{'FundCode'};

    $values->{'PerSolicitorId'} = '';
    my($nameMatchCount) = $dbh->selectrow_array(q{
      select count(*)
        from name_map
        where c_name = ?
      }, undef, $values->{'VolunteerName'});
    
    if ($nameMatchCount == 1) {
      ($values->{'PerSolicitorId'}) = $dbh->selectrow_array(q{
        select p_id
          from name_map
          where c_name = ?
        }, undef, $values->{'VolunteerName'});
    }

    unless ($values->{'ProductCode'}) {
      write_record($noProductCodeWorksheet, $noProductRow++, [
        'program',
        $values->{'BranchName'},
        '',
        '',
        $values->{'ItemDescription'},
        '',
        ''
      ]);

      return;    
    }

    my $record = make_record($values, \@allColumns, $columnMap);
    write_record($worksheet, $order++, $record);

    $record = make_record($values, [donation_order_fields()], make_column_map([donation_order_fields()]));
    $csv->print($donationMaster, $record);
    $csv->print($allOrders, [
      $values->{'OrderNo'}, 
      lookup_t_id($values->{'PerMemberId'}),
      $values->{'PerMemberId'},
      lookup_t_id($values->{'PerBillableMemberId'}),
      $values->{'PerBillableMemberId'},
      'pledge'
    ]);

  },
  undef,
  $orderHeaderMap
);

close($donationMaster);

$orderHeaderMap = {
  'Amount' => 'OrderTotal',
  'Amount Due' => 'AmountDue',
  'Branch' => 'BranchName',
  'Date' => 'OrderDate',
  'Days Past Due' => 'DaysPastDue',
  'Description' => 'ItemDescription',
  'Do Not Mail Invoice' => 'DoNotMailInvoice',
  'End Date' => 'EndDate',
  'GL#' => 'GlAccount',
  'Member ID' => 'MemberId',
  'Membership Status' => 'MembershipStatus',
  'Membership Type' => 'MembershipType',
  'Notes' => 'Notes',
  'Reference #' => 'ReceiptNumber',
  'Sales Tax' => 'SalesTax',
  'Start Date' => 'StartDate',
  'Total Due' => 'BalanceDue',
  'Total Pledge Amount Due' => 'TotalPledgeAmountDue',
  'Type' => 'Comments',
};

open(my $arBalMaster, '>', 'data/arbal_orders.csv')
  or die "Couldn't open data/arbal_orders.csv: $!";
$csv->print($arBalMaster, [arbal_order_fields()]);

foreach my $arFile (qw( Camps Childcare Counter Memberships Programs)) {
  process_data_file(
    'data/AR' . $arFile . '.csv',
    sub {
      my $values = shift;
      # dd($values); 

      return if (Date_Cmp($currentDate, ParseDate($values->{'OrderDate'})) == -1);
      return if ($values->{'MemberId'} =~ /^\d$/);

      $values->{'OrderNo'} = $orderNo++;

      $values->{'ProductCode'} = 'AR_BAL_CONV';
      $values->{'DatePaid'} = '';

      $values->{'StatusDate'} = $values->{'OrderDate'};

      $values->{'PerMemberId'} = lookup_id($values->{'MemberId'});
      $values->{'PerBillableMemberId'} = $values->{'PerMemberId'};

      $values->{'OrderTotal'} =~ s/[\$,]//g;
      $values->{'BalanceDue'} =~ s/[\$,]//g;

      my $record = make_record($values, \@allColumns, $columnMap);
      write_record($worksheet, $order++, $record);

      $record = make_record($values, [arbal_order_fields()], make_column_map([arbal_order_fields()]));
      $csv->print($arBalMaster, $record);
      $csv->print($allOrders, [
        $values->{'OrderNo'}, 
        lookup_t_id($values->{'PerMemberId'}),
        $values->{'PerMemberId'},
        lookup_t_id($values->{'PerBillableMemberId'}),
        $values->{'PerBillableMemberId'},
        'ar'
      ]);

    },
    undef,
    $orderHeaderMap
  );
}
close($arBalMaster);

sub lookup_product_code {
  my $type = shift;
  my $values = shift;

  my $branchName = $values->{'BranchName'};
  my $description = $values->{'ProgramDescription'};
  my $summary = $values->{'ItemDescription'};
  my $session = $values->{'Session'};

  # $description =~ s/ +/ /g;
  # $summary =~ s/ +/ /g;

  my $query = q{
    select product_code
      from products
      where branch = ?
        and description = ?
        and summary = ?
    };

  $query .= q{
        and session = ?
      } if ($type eq 'camp');

  if ($values->{'Cycle'} =~ /2018 Summer (\d)/) {
    $query .= qq{
          and product_code like '%_M${1}8'
        };
  }

  my $sth = $dbh->prepare($query);

  $sth->bind_param(1, $branchName);
  $sth->bind_param(2, $description);
  $sth->bind_param(3, $summary);
  $sth->bind_param(4, $session) if ($type eq 'camp');

  $sth->execute();

  my($productCode) = $sth->fetchrow_array();

  return $productCode;
}

sub lookup_campaign_product {
  my $values = shift;

  my $details = {
    'ProductCode' => '',
    'BranchName' => '',
    'BranchCode' => '',
    'CampaignCode' => '',
    'FundCode' => 'ANNUAL_FUND'
  };

  $details->{'BranchName'} = resolve_branch_name($values);
  $details->{'BranchCode'} = branch_name_map()->{$details->{'BranchName'}};
  
  if ($values->{'ItemDescription'} =~ /(20\d{2})/) {
    $details->{'CampaignCode'} = $1 . '_' . $details->{'BranchCode'};
  }

  my $eventType = 'GN';
  $eventType = 'EVENT' if ($values->{'ItemDescription'} =~ /event/i);

  my $fundType = ($values->{'BalanceDue'} > 0) ? 'PL' : 'CA';

  $details->{'ProductCode'} = join('_', $details->{'BranchCode'}, 'ANNUAL', $eventType, $fundType);

  return $details;
}

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

sub all_order_fields {
  return qw(
    order_number
    trinexum_id
    personify_id
    billable_trx_id
    billable_per_id
    type
  );
}