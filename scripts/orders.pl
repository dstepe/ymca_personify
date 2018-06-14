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
) if (0);

close($orderMaster);

print Dumper($unmappedMembers) if (@{$unmappedMembers});

my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

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

    $values->{'Balance'} = '';
    $values->{'ProgramFee'} = '';

    $values->{'FeePaid'} =~ s/[\$,]//g;

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
  },
  undef,
  $orderHeaderMap
) if (0);

$orderHeaderMap = {
  'Branch' => 'BranchName',
  'Class Summary' => 'ItemDescription',
  'Current Balance' => 'Balance',
  'Date Enrolled' => 'DateEnrolled',
  'Participant Id' => 'ParticipantId',
  'Primary Sponsor Id' => 'PrimarySponsorId',
  'Program Description' => 'ProgramDescription',
  'Program Fee' => 'ProgramFee',
  'Session' => 'Session',
  'Start Date' => 'ProgramStartDate',
};

foreach my $branchFile (qw( Atrium EastButler Fairfield Fitton Middletown )) {
  process_data_file(
    'data/Camp' . $branchFile . '.csv',
    sub {
      my $values = shift;
      # dd($values);

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

      $values->{'Balance'} =~ s/[\$,]//g;
      $values->{'ProgramFee'} =~ s/[\$,]//g;

      $values->{'FeePaid'} = $values->{'ProgramFee'} - $values->{'Balance'};

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
    },
    undef,
    $orderHeaderMap
  ) if (0);
}

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
    $values->{'ProgramFee'} = '';
    $values->{'FeePaid'} = $values->{'CampaignPledge'};
    $values->{'Balance'} = $values->{'CampaignBalance'};
    $values->{'DatePaid'} = UnixDate($values->{'PledgeDate'}, '%Y-%m-%d');
    $values->{'Cycle'} = '';

    $values->{'PerMemberId'} = lookup_id($values->{'MemberId'});
    $values->{'PerBillableMemberId'} = $values->{'PerMemberId'};

    $values->{'OrderDate'} = $values->{'PledgeDate'};
    $values->{'StatusDate'} = $values->{'PledgeDate'};

    $values->{'FeePaid'} =~ s/[\$,]//g;
    $values->{'Balance'} =~ s/[\$,]//g;

    $values->{'Comments'} =~ s/^\s+//;
    $values->{'Comments'} =~ s/\s+$//;
    $values->{'Comments'} = encode_base64($values->{'Comments'}, '');

    my $campaignProductDetails = lookup_campaign_product($values);

    $values->{'BranchName'} = $campaignProductDetails->{'BranchName'};
    $values->{'BranchCode'} = $campaignProductDetails->{'BranchCode'};
    $values->{'ProductCode'} = $campaignProductDetails->{'ProductCode'};
    $values->{'CampaignCode'} = $campaignProductDetails->{'CampaignCode'};
    $values->{'FundCode'} = $campaignProductDetails->{'FundCode'};
    # dd($values);
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
  },
  undef,
  $orderHeaderMap
);

close($donationMaster);

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

  my $fundType = ($values->{'Balance'} > 0) ? 'PL' : 'CA';

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