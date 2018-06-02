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

my $templateName = 'DCT_ORDER_DETAIL-32358';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ORDER_DATE'                         => { 'type' => 'record', 'source' => 'OrderDate' },
  'SHIP_CUSTOMER_ID'                   => { 'type' => 'record', 'source' => 'ShipCustomerId' },
  'SHIP_ADDRESS_TYPE_CODE'             => { 'type' => 'static', 'source' => 'HOME' },
  'INVOICE_NO'                         => { 'type' => 'record', 'source' => 'TrxInvoiceId' },
  'INVOICE_DATE'                       => { 'type' => 'record', 'source' => 'OrderDate' },
  'SUBSYSTEM'                          => { 'type' => 'static', 'source' => 'MBR' },
  'PRODUCT_CODE'                       => { 'type' => 'record', 'source' => 'ProductCode' },
  'PARENT_PRODUCT'                     => { 'type' => 'static', 'source' => 'GMV' },
  'LINE_TYPE'                          => { 'type' => 'static', 'source' => 'IP' },
  'LINE_STATUS_CODE'                   => { 'type' => 'static', 'source' => 'A' },
  'LINE_STATUS_DATE'                   => { 'type' => 'record', 'source' => 'OrderDate' },
  'FULFILL_STATUS_CODE'                => { 'type' => 'static', 'source' => 'A' },
  'FULFILL_STATUS_DATE'                => { 'type' => 'record', 'source' => 'OrderDate' },
  'RECOGNITION_STATUS_CODE'            => { 'type' => 'static', 'source' => 'C' },
  'RATE_STRUCTURE'                     => { 'type' => 'static', 'source' => 'LIST' },
  'RATE_CODE'                          => { 'type' => 'record', 'source' => 'RateCode' },
  'TAXABLE_FLAG'                       => { 'type' => 'static', 'source' => 'Y' },
  'TAX_CATEGORY_CODE'                  => { 'type' => 'static', 'source' => 'SALES' },
  'REQUESTED_QTY'                      => { 'type' => 'static', 'source' => '1' },
  'ORDER_QTY'                          => { 'type' => 'static', 'source' => '1' },
  'TOTAL_AMOUNT'                       => { 'type' => 'record', 'source' => 'TotalAmount' },
  'CYCLE_BEGIN_DATE'                   => { 'type' => 'record', 'source' => 'BeginDate' },
  'CYCLE_END_DATE'                     => { 'type' => 'record', 'source' => 'EndDate' },
  'BACK_ISSUE_FLAG'                    => { 'type' => 'static', 'source' => 'Y' },
  'INITIAL_BEGIN_DATE'                 => { 'type' => 'record', 'source' => 'JoinDate' },
  'RETURNED_QTY'                       => { 'type' => 'static', 'source' => '0' },
  'PAYOR_CUSTOMER_ID'                  => { 'type' => 'record', 'source' => 'BillCustomerId' },
  'RECEIPT_TYPE'                       => { 'type' => 'static', 'source' => 'CASH' },
  'RECEIPT_CURRENCY_CODE'              => { 'type' => 'static', 'source' => 'USD' },
  'RECEIPT_DATE'                       => { 'type' => 'record', 'source' => 'OrderDate' },
  'XRATE'                              => { 'type' => 'static', 'source' => '1' },
  'PAYMENT_AMOUNT'                     => { 'type' => 'record', 'source' => 'TotalAmount' },
  'RECEIPT_STATUS_CODE'                => { 'type' => 'static', 'source' => 'A' },
  'RECEIPT_STATUS_DATE'                => { 'type' => 'record', 'source' => 'OrderDate' },
  'CL_LATE_FEE_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'MANUAL_DISCOUNT_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'DISCOUNT_CODE'                      => { 'type' => 'record', 'source' => 'DiscountCode' },
  'ACTUAL_DISCOUNT_AMOUNT'             => { 'type' => 'record', 'source' => 'DiscountAmount' },
  'ACTUAL_SHIP_AMOUNT'                 => { 'type' => 'static', 'source' => '0' },
  'ACTUAL_TAX_AMOUNT'                  => { 'type' => 'record', 'source' => 'TaxPaidAmount' },
  'MARKET_CODE'                        => { 'type' => 'record', 'source' => 'MarketCode' },
  'COMMENTS_ON_INVOICE_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'AUTO_PAY_METHOD_CODE'               => { 'type' => 'static', 'source' => 'NONE' },
  'ATTENDANCE_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'BLOCK_SALES_TAX_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'LINE_COMPLETE_FLAG'                 => { 'type' => 'static', 'source' => 'N' },
  'RENEW_TO_CC_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'RENEWAL_CREATED_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'REQUIRES_DISCOUNT_CALCULATION_FLAG' => { 'type' => 'static', 'source' => 'Y' },
  'TOTAL_DEFERRED_TAX'                 => { 'type' => 'static', 'source' => '0' },
  'TOTAL_DEPOSIT_TAX'                  => { 'type' => 'static', 'source' => '0' },
};

my $taxRates = {
  'Atrium' => .07,
  'Other' => .065,
};

my $branchProgramCodes = branch_name_map();

my $cycleDurations = {
  'Annual' => '1 year',
  'Monthly E-Pay' => '1 month',
  'Quarterly' => '3 months',
};

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

my $prdRates = {};
my($rateFile, $headers, $totalRows) = open_data_file('data/PrdRates.csv');
while(my $rowIn = $csv->getline($rateFile)) {
  my $values = map_values($headers, $rowIn);

  my $type = uc $values->{'Current Type'};
  $prdRates->{$type} = {
    'NewType' => $values->{'New Type'},
  };

  ($prdRates->{$type}{'Monthly'} = $values->{'Monthly Amt'}) =~ s/\$//g;
  ($prdRates->{$type}{'Annual'} = $values->{'Annual Amt'}) =~ s/\$//g;
}
close($rateFile);

my $membershipMap = {};
my $mappingFile;
($mappingFile, $headers, $totalRows) = open_data_file('data/MembershipMapping.csv');
while(my $rowIn = $csv->getline($mappingFile)) {
  my $values = map_values($headers, $rowIn);

  $membershipMap->{uc $values->{'Description'}}{uc $values->{'Current PaymentMethod'}} = {
    'Branch' => $values->{'Branch'},
    'PaymentMethod' => $values->{'Current PaymentMethod'},
    'ProductCode' => $values->{'Membership Product Code'},
    'RateCode' => $values->{'Rate Code'},
    'MarketCode' => $values->{'Market Code'},
    'DiscountCode' => $values->{'Discount Code'},
    'DiscountAmount' => $values->{'Discount Amount'},
    'PurchasingGroup' => $values->{'Purchasing Group'},
    'Discount' => $values->{'Discount'},
  };
}
close($mappingFile);

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $missingMembershipMap = {};
my $row = 1;
process_data_file(
  'data/member_orders.csv',
  sub {
    my $values = shift;

    my $membershipTypeKey = uc $values->{'MembershipType'};
    my $paymentMethodKey = uc $values->{'PaymentMethod'};

    my $taxRate = $taxRates->{'Other'};

    if (exists($taxRates->{$values->{'MembershipBranch'}})) {
      $taxRate = $taxRates->{$values->{'MembershipBranch'}};
    }

    my $orderDate = ParseDate($values->{'OrderDate'});
    $values->{'BeginDate'} = UnixDate($orderDate, '%Y-%m-%d');
    my $cycle = $cycleDurations->{$values->{'PaymentMethod'}};
    $values->{'EndDate'} = UnixDate(DateCalc(DateCalc($orderDate, '+' . $cycle), '-1 day'), '%Y-%m-%d');

    $values->{'TrxInvoiceId'} = '';
    $values->{'ProductCode'} = '';
    $values->{'RateCode'} = '';
    $values->{'DiscountAmount'} = 0;
    $values->{'DiscountCode'} = '';

    my $discount = '';
    if (exists($membershipMap->{$membershipTypeKey}) 
        && exists($membershipMap->{$membershipTypeKey}{$paymentMethodKey})) {
      my $map = $membershipMap->{$membershipTypeKey}{$paymentMethodKey};

      my $branchCode = $branchProgramCodes->{$values->{'MembershipBranch'}};
      
      $values->{'ProductCode'} = $map->{'ProductCode'};
      $values->{'RateCode'} = $map->{'RateCode'};
      $values->{'MarketCode'} = $map->{'MarketCode'};
      $values->{'DiscountCode'} = $map->{'DiscountCode'};
      if (!$map->{'Branch'}) {
        $values->{'ProductCode'} =~ s/\{BRANCH\}/$branchCode/g;
        $values->{'MarketCode'} =~ s/\{BRANCH\}/$branchCode/g;
        $values->{'DiscountCode'} =~ s/\{BRANCH\}/$branchCode/g;
      }

      $discount = $map->{'DiscountAmount'};
    }
    
    if ($membershipTypeKey =~ /SPONSOR/i) {
      $values->{'DiscountCode'} = 'SPONSOR' . $values->{'SponsorDiscount'};
      $discount = $values->{'SponsorDiscount'} . '%';
    }

    $missingMembershipMap->{$membershipTypeKey}{$paymentMethodKey}++ unless ($values->{'RateCode'});

    my $baseFee = $values->{'RenewalFee'};

    if ($membershipTypeKey =~ /PRD/) {
      (my $prd = $membershipTypeKey) =~ s/\-.*//;
      die "Missing PRD mapping for $prd {$values->{'RateCode'}" 
        unless (exists($prdRates->{$prd}{$values->{'RateCode'}}));
      $baseFee = $prdRates->{$prd}{$values->{'RateCode'}};
    }

    $values->{'DiscountAmount'} = 0;
    if ($discount =~ /\%/) {
      $discount =~ s/\%//;
      $discount /= 100;

      $values->{'DiscountAmount'} = $baseFee * $discount;
    } elsif ($discount =~ /\$/) {
      $discount =~ s/\$//;

      $values->{'DiscountAmount'} = $discount;
    }

    my $finalFee = $baseFee - $values->{'DiscountAmount'};
    $finalFee = 0 if ($finalFee < 0);

    $values->{'TaxPaidAmount'} = sprintf("%.2f", $finalFee * $taxRate);

    $values->{'TotalAmount'} = $finalFee + $values->{'TaxPaidAmount'};

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);

print Dumper($missingMembershipMap) if (keys %{$missingMembershipMap});
