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

my $templateName = 'Order_Detail_CC_Info';

my $columnMap = {
  'ORDER_NO'                   => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'              => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_REFERENCE'               => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_ISSUER_NO'               => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_START_DATE'              => { 'type' => 'record', 'source' => 'OrderNo' },
  'EXP_DATE'                   => { 'type' => 'record', 'source' => 'OrderNo' },
  'RECEIPT_TYPE_CODE'          => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_NAME'                    => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_ADDRESS_1'               => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_CITY'                    => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_STATE'                   => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_POSTAL_CODE'             => { 'type' => 'record', 'source' => 'OrderNo' },
  'CC_COUNTRY_CODE'            => { 'type' => 'record', 'source' => 'OrderNo' },
  'MERCHANT_ID'                => { 'type' => 'record', 'source' => 'OrderNo' },
  'PAYMENT_HANDLER_CODE'       => { 'type' => 'record', 'source' => 'OrderNo' },
  'CUS_CREDIT_CARD_PROFILE_ID' => { 'type' => 'record', 'source' => 'OrderNo' },
  'PERSONIFY_ORDER_NO'         => { 'type' => 'record', 'source' => 'OrderNo' },
};

# load all order information
# while processing cc rows, see if member has a non-membership order
# collect all cc rows and associated orders
# use separate loop to write unique cc records with multiple orders

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my $orders = {};
my($orderFile, $headers, $totalRows) = open_data_file('data/all_orders.csv');
while(my $rowIn = $csv->getline($orderFile)) {
  my $values = map_values($headers, $rowIn);
  # dd($values);

  $orders->{$values->{'trinexum_id'}} = [] unless (exists($orders->{$values->{'trinexum_id'}}));
  push(@{$orders->{$values->{'trinexum_id'}}}, $values);
}
close($orderFile);

foreach my $trxId (keys %{$orders}) {
  next unless (scalar(@{$orders->{$trxId}}) > 2);
  print "$trxId has orders exceeding 2: " . scalar(@{$orders->{$trxId}}) . "\n";
}

exit;

my $row = 1;
process_data_file(
  'data/cc.csv',
  sub {
    my $values = shift;
    dd($values);

    
    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);

