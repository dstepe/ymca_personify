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
use MIME::Base64;
use Encode;

my $templateName = 'Cus_EFT_Info_1';

my $columnMap = {
  'CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'PerMemberId' },
  'TRX_ID'                   => { 'type' => 'record', 'source' => 'TrxMemberId' },
  'ROUTING_NUMBER'           => { 'type' => 'record', 'source' => 'RoutingNumber' },
  'BANK_ACCOUNT_NUMBER'      => { 'type' => 'record', 'source' => 'AccountNumber' },
  'ACCOUNT_DESCR'            => { 'type' => 'record', 'source' => 'AccountDesc' },
  'BANK_NAME'                => { 'type' => 'record', 'source' => 'BankName' },
  'ACCOUNT_TYPE_CODE'        => { 'type' => 'record', 'source' => 'AccountTypeCode' },
  'BEGIN_DATE'               => { 'type' => 'record', 'source' => 'BeginDate' },
  'BANK_ACCOUNT_STATUS_CODE' => { 'type' => 'static', 'source' => 'GOOD' },
  'BANK_ACCOUNT_STATUS_DATE' => { 'type' => 'record', 'source' => 'StatusDate' },
  'PERSONIFY_ORDER_NO_1'     => { 'type' => 'record', 'source' => 'OrderNo1' },
  'PERSONIFY_ORDER_NO_2'     => { 'type' => 'record', 'source' => 'OrderNo2' },
  'PERSONIFY_ORDER_NO_3'     => { 'type' => 'record', 'source' => 'OrderNo3' },
};

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my $orders = {};
my($orderFile, $headers, $totalRows) = open_data_file('data/all_orders.csv');
while(my $rowIn = $csv->getline($orderFile)) {
  my $values = map_values($headers, $rowIn);
  # dd($values);

  next unless ($values->{'type'} eq 'membership');
  $orders->{$values->{'trinexum_id'}} = [] unless (exists($orders->{$values->{'trinexum_id'}}));
  push(@{$orders->{$values->{'trinexum_id'}}}, $values);
}
close($orderFile);

foreach my $trxId (keys %{$orders}) {
  next unless (scalar(@{$orders->{$trxId}}) > 2);
  print "$trxId has orders exceeding 2: " . scalar(@{$orders->{$trxId}}) . "\n";
}

my $row = 1;
process_data_file(
  'data/eft.csv',
  sub {
    my $values = shift;
    # dd($values);

    $values->{'AccountDesc'} = $values->{'ACCOUNT_DESCR'};
    $values->{'AccountTypeCode'} = $values->{'ACCOUNT_TYPE_CODE'};
    $values->{'AccountNumber'} = $values->{'BANK_ACCOUNT_NUMBER'};
    $values->{'BankName'} = $values->{'BANK_NAME'};
    $values->{'RoutingNumber'} = $values->{'ROUTING_NUMBER'};
    $values->{'TrxMemberId'} = $values->{'TRX_ID'};

    $values->{'PerMemberId'} = lookup_id($values->{'TRX_ID'});
    $values->{'BeginDate'} = '2017-01-01';
    $values->{'StatusDate'} = '2017-01-01';
    
    $values->{'OrderNo1'} = '';
    $values->{'OrderNo2'} = '';
    $values->{'OrderNo3'} = '';

    if (exists($orders->{$values->{'TRX_ID'}})) {
      my $memberOrders = $orders->{$values->{'TRX_ID'}};
      my $orderCount = scalar(@{$memberOrders});

      for (my $i = 0; $i < $orderCount; $i++) {
        $values->{'OrderNo' . ($i + 1)} = $memberOrders->[$i]{'order_number'};
      }
    }

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);

