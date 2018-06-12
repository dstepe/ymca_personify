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

my $templateName = 'DCT_ORDER_DETAIL-32358';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ORDER_DATE'                         => { 'type' => 'record', 'source' => 'OrderDate' },
  'SHIP_CUSTOMER_ID'                   => { 'type' => 'record', 'source' => 'ShipCustomerId' },
  'SHIP_ADDRESS_TYPE_CODE'             => { 'type' => 'static', 'source' => 'HOME' },
  'INVOICE_NO'                         => { 'type' => 'record', 'source' => 'TrxInvoiceId' },
  'INVOICE_DATE'                       => { 'type' => 'record', 'source' => 'OrderDate' },
  'SUBSYSTEM'                          => { 'type' => 'record', 'source' => 'SubSystem' },
  'PRODUCT_CODE'                       => { 'type' => 'record', 'source' => 'ProductCode' },
  'PARENT_PRODUCT'                     => { 'type' => 'record', 'source' => 'ParentProductCode' },
  'LINE_TYPE'                          => { 'type' => 'static', 'source' => 'IP' },
  'LINE_STATUS_CODE'                   => { 'type' => 'static', 'source' => 'A' },
  'LINE_STATUS_DATE'                   => { 'type' => 'record', 'source' => 'OrderDate' },
  'FULFILL_STATUS_CODE'                => { 'type' => 'static', 'source' => 'A' },
  'FULFILL_STATUS_DATE'                => { 'type' => 'record', 'source' => 'FullfillStatusDate' },
  'RECOGNITION_STATUS_CODE'            => { 'type' => 'static', 'source' => 'C' },
  'RATE_STRUCTURE'                     => { 'type' => 'static', 'source' => 'LIST' },
  'RATE_CODE'                          => { 'type' => 'record', 'source' => 'RateCode' },
  'TAXABLE_FLAG'                       => { 'type' => 'record', 'source' => 'TaxableFlag' },
  'TAX_CATEGORY_CODE'                  => { 'type' => 'record', 'source' => 'TaxCategoryCode' },
  'REQUESTED_QTY'                      => { 'type' => 'static', 'source' => '1' },
  'ORDER_QTY'                          => { 'type' => 'static', 'source' => '1' },
  'TOTAL_AMOUNT'                       => { 'type' => 'record', 'source' => 'TotalAmount' },
  'CYCLE_BEGIN_DATE'                   => { 'type' => 'record', 'source' => 'BeginDate' },
  'CYCLE_END_DATE'                     => { 'type' => 'record', 'source' => 'EndDate' },
  'BACK_ISSUE_FLAG'                    => { 'type' => 'record', 'source' => 'BackIssueFlag' },
  'INITIAL_BEGIN_DATE'                 => { 'type' => 'record', 'source' => 'JoinDate' },
  'DUE_DATE'                           => { 'type' => 'record', 'source' => 'DueDate' },
  'RETURNED_QTY'                       => { 'type' => 'static', 'source' => '0' },
  'PAY_FREQUENCY_CODE'                 => { 'type' => 'record', 'source' => 'PayFrequencyCode' },
  'PAYOR_CUSTOMER_ID'                  => { 'type' => 'record', 'source' => 'PerBillableMemberId' },
  'RECEIPT_TYPE'                       => { 'type' => 'static', 'source' => 'CASH' },
  'RECEIPT_CURRENCY_CODE'              => { 'type' => 'static', 'source' => 'USD' },
  'RECEIPT_DATE'                       => { 'type' => 'record', 'source' => 'OrderDate' },
  'XRATE'                              => { 'type' => 'static', 'source' => '1' },
  'PAYMENT_AMOUNT'                     => { 'type' => 'record', 'source' => 'TotalAmount' },
  'RECEIPT_STATUS_CODE'                => { 'type' => 'static', 'source' => 'A' },
  'RECEIPT_STATUS_DATE'                => { 'type' => 'record', 'source' => 'OrderDate' },
  'CL_LATE_FEE_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'PRICING_CURRENCY_CODE'              => { 'type' => 'record', 'source' => 'PricingDiscountCode' },
  'MANUAL_DISCOUNT_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'DISCOUNT_CODE'                      => { 'type' => 'record', 'source' => 'DiscountCode' },
  'ACTUAL_DISCOUNT_AMOUNT'             => { 'type' => 'record', 'source' => 'DiscountAmount' },
  'ACTUAL_SHIP_AMOUNT'                 => { 'type' => 'static', 'source' => '0' },
  'ACTUAL_TAX_AMOUNT'                  => { 'type' => 'record', 'source' => 'TaxPaidAmount' },
  'MARKET_CODE'                        => { 'type' => 'record', 'source' => 'MarketCode' },
  'CAMPAIGN'                           => { 'type' => 'record', 'source' => 'Campaign' },
  'FUND'                               => { 'type' => 'record', 'source' => 'Fund' },
  'APPEAL'                             => { 'type' => 'record', 'source' => 'Appeal' },
  'COMMENTS_ON_INVOICE_FLAG'           => { 'type' => 'record', 'source' => 'CommentsOnInvoice' },
  'DESCRIPTION'                        => { 'type' => 'record', 'source' => 'InvoiceDescription' },
  'AUTO_PAY_METHOD_CODE'               => { 'type' => 'static', 'source' => 'NONE' },
  'ATTENDANCE_FLAG'                    => { 'type' => 'record', 'source' => 'AttendanceFlag' },
  'BLOCK_SALES_TAX_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'LINE_COMPLETE_FLAG'                 => { 'type' => 'static', 'source' => 'N' },
  'RENEW_TO_CC_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'RENEWAL_CREATED_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'REQUIRES_DISCOUNT_CALCULATION_FLAG' => { 'type' => 'record', 'source' => 'RequireDiscountCalc' },
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

    $values->{'SubSystem'} = 'MBR';
    $values->{'ParentProductCode'} = 'GMV';
    $values->{'TaxableFlag'} = 'Y';
    $values->{'TaxCategoryCode'} = 'SALES';
    $values->{'AttendanceFlag'} = 'N';
    $values->{'PayFrequencyCode'} = '';
    $values->{'RequireDiscountCalc'} = 'Y';
    $values->{'BackIssueFlag'} = 'Y';
    $values->{'DueDate'} = '';
    $values->{'PricingDiscountCode'} = '';
    $values->{'Campaign'} = '';
    $values->{'Fund'} = '';
    $values->{'Appeal'} = '';
    $values->{'CommentsOnInvoice'} = '';
    $values->{'InvoiceDescription'} = '';

    $values->{'ShipCustomerId'} = $values->{'PerBillableMemberId'};

    my $membershipTypeKey = uc $values->{'MembershipTypeDes'};
    my $paymentMethodKey = uc $values->{'PaymentMethod'};

    my $taxRate = $taxRates->{'Other'};

    if (exists($taxRates->{$values->{'MembershipBranch'}})) {
      $taxRate = $taxRates->{$values->{'MembershipBranch'}};
    }

    $values->{'FullfillStatusDate'} = $values->{'OrderDate'};

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

    my $baseFee = $values->{'RenewMembershipFee'};

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

process_data_file(
  'data/program_orders.csv',
  sub {
    my $values = shift;
    # dd($values);

    $values->{'SubSystem'} = 'MTG';
    $values->{'RateCode'} = 'STD';
    $values->{'MarketCode'} = '';
    $values->{'TaxableFlag'} = 'N';
    $values->{'TaxCategoryCode'} = '';
    $values->{'DiscountAmount'} = 0;
    $values->{'DiscountCode'} = '';
    $values->{'TaxPaidAmount'} = 0;
    $values->{'AttendanceFlag'} = 'Y';
    $values->{'JoinDate'} = '';
    $values->{'PayFrequencyCode'} = '';
    $values->{'RequireDiscountCalc'} = 'Y';
    $values->{'BackIssueFlag'} = 'Y';
    $values->{'DueDate'} = '';
    $values->{'PricingDiscountCode'} = '';
    $values->{'Campaign'} = '';
    $values->{'Fund'} = '';
    $values->{'Appeal'} = '';
    $values->{'CommentsOnInvoice'} = '';
    $values->{'InvoiceDescription'} = '';
    
    $values->{'ShipCustomerId'} = $values->{'PerMemberId'};

    $values->{'BeginDate'} = UnixDate($values->{'ProgramStartDate'}, '%Y-%m-%d');
    $values->{'EndDate'} = UnixDate($values->{'ProgramEndDate'}, '%Y-%m-%d');

    $values->{'TrxInvoiceId'} = $values->{'ReceiptNumber'};

    $values->{'TotalAmount'} = $values->{'FeePaid'};

    $values->{'ParentProductCode'} = $values->{'ProductCode'};

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);

my $today = ParseDate('today');

process_data_file(
  'data/donation_orders.csv',
  sub {
    my $values = shift;
    # dd($values);

    $values->{'SubSystem'} = 'FND';
    $values->{'RateCode'} = 'STD';
    $values->{'MarketCode'} = '';
    $values->{'TaxableFlag'} = 'N';
    $values->{'TaxCategoryCode'} = '';
    $values->{'DiscountAmount'} = 0;
    $values->{'DiscountCode'} = '';
    $values->{'TaxPaidAmount'} = 0;
    $values->{'AttendanceFlag'} = 'N';
    $values->{'JoinDate'} = '';
    $values->{'PayFrequencyCode'} = 'IMMEDIATE';
    $values->{'RequireDiscountCalc'} = 'N';
    $values->{'BackIssueFlag'} = '';
    $values->{'PricingDiscountCode'} = 'USD';
    $values->{'Campaign'} = '';
    $values->{'Fund'} = '';
    $values->{'Appeal'} = '';
    $values->{'InvoiceDescription'} = '';
    
    $values->{'FullfillStatusDate'} = '';

    $values->{'Comments'} = decode_base64($values->{'Comments'});
    $values->{'CommentsOnInvoice'} = 'N';
    $values->{'CommentsOnInvoice'} = 'Y' if ($values->{'Comments'});

    $values->{'ShipCustomerId'} = $values->{'PerMemberId'};

    $values->{'BeginDate'} = '';
    $values->{'EndDate'} = '';

    $values->{'DueDate'} = '';
    my $nextBillDate = ParseDate($values->{'PledgeNextBillDate'});
    if (Date_Cmp($today, $nextBillDate) <= 0) {
      $values->{'DueDate'} = UnixDate($nextBillDate, '%Y-%m-%d');
    }

    $values->{'TrxInvoiceId'} = $values->{'ReceiptNumber'};

    $values->{'TotalAmount'} = $values->{'FeePaid'};

    $values->{'ParentProductCode'} = $values->{'ProductCode'};

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);
