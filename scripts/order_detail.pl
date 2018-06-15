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
  'PAYMENT_AMOUNT'                     => { 'type' => 'record', 'source' => 'PaymentAmount' },
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
  'CAMPAIGN'                           => { 'type' => 'record', 'source' => 'CampaignCode' },
  'FUND'                               => { 'type' => 'record', 'source' => 'FundCode' },
  'APPEAL'                             => { 'type' => 'record', 'source' => 'AppealCode' },
  'COMMENTS_ON_INVOICE_FLAG'           => { 'type' => 'record', 'source' => 'CommentsOnInvoice' },
  'DESCRIPTION'                        => { 'type' => 'record', 'source' => 'InvoiceDescription' },
  'COMMENTS'                           => { 'type' => 'record', 'source' => 'Comments' },
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

my $fndOrderTemplateName = 'DCT_FND_ORDER_DETAIL-67574';

my $fndOrderColumnMap = {
  'ORDER_NO'                               => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'                          => { 'type' => 'static', 'source' => '1' },
  'LIST_DONOR_AS'                          => { 'type' => 'record', 'source' => 'DonorName' },
  'SOFT_CREDIT_MAST_CUST'                  => { 'type' => 'record', 'source' => 'SoftCreditCustId' },
  'RESTRICTED_GIFT_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'CREATE_TRANSACTION_FLAG'                => { 'type' => 'static', 'source' => 'Y' },
  'RECOGNIZED_FLAG'                        => { 'type' => 'static', 'source' => 'Y' },
  'TRIBUTE_TYPE_CODE'                      => { 'type' => 'record', 'source' => 'TributeTypeCode' },
  'IN_TRIBUTE_TO_MAST_CUST'                => { 'type' => 'record', 'source' => 'InTributeToCustId' },
  'ANONYMOUS_FLAG'                         => { 'type' => 'record', 'source' => 'Anonymous' },
  'GIVE_TRIBUTE_CUSTOMER_CREDIT_FLAG'      => { 'type' => 'static', 'source' => 'N' },
  'IN_TRIBUTE_TO_LABEL_NAME'               => { 'type' => 'record', 'source' => 'InTributeToCustName' },
  'RECURRING_GIFT_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'COMPANY_MATCHES_GIFT_FLAG'              => { 'type' => 'static', 'source' => 'N' },
};

my $fndHardCreditTemplateName = 'DCT_FND_HARD_CREDIT-95863';

my $fndHardCreditColumnMap = {
  'ORG_ID'                => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'           => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'CUSTOMER_ID'           => { 'type' => 'record', 'source' => 'PerMemberId' },
  'TXN_DATE'              => { 'type' => 'record', 'source' => 'DatePaid' },
  'PAYMENT_BASED_FLAG'    => { 'type' => 'static', 'source' => 'N' },
  'PARENT_PRODUCT'        => { 'type' => 'record', 'source' => 'ParentProductCode' },
  'PRODUCT_CODE'          => { 'type' => 'record', 'source' => 'ProductCode' },
  'ORDER_NO'              => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'         => { 'type' => 'static', 'source' => '1' },
  'CAMPAIGN'              => { 'type' => 'record', 'source' => 'CampaignCode' },
  'FUND'                  => { 'type' => 'record', 'source' => 'FundCode' },
  'APPEAL'                => { 'type' => 'record', 'source' => 'AppealCode' },
  'SOLICITOR_CUSTOMER_ID' => { 'type' => 'record', 'source' => 'PerSolicitorId' },
  'CREDIT_AMOUNT'         => { 'type' => 'record', 'source' => 'TotalAmount' },
  'CREDIT_TYPE_CODE'      => { 'type' => 'static', 'source' => 'Bill' },
  'ACK_LETTER_CODE'       => { 'type' => 'static', 'source' => 'HISTORY' },
  'COMMENTS'              => { 'type' => 'record', 'source' => 'Comments' },
};

my $taxRates = {
  'Atrium' => .07,
  'Other' => .065,
};

my $branchProgramCodes = branch_name_map();

my $cycleDurations = {
  'Annual' => '1 year',
  'Monthly E-Pay' => '1 month',
  'Monthly' => '1 month',
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

  ($prdRates->{$type}{'Monthly'} = $values->{'Monthly Amt'}) =~ s/[\$,]//g;
  ($prdRates->{$type}{'Annual'} = $values->{'Annual Amt'}) =~ s/[\$,]//g;
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

my $orderProblemsWorkbook = make_workbook('order_problems');
our $orderProblemsWorksheet = make_worksheet($orderProblemsWorkbook, 
  ['Source', 'Bill Id', 'Description', 'Problem']);
our $orderProblemRow = 1;

my $denyEndDate = UnixDate('12-31-2099', '%Y-%m-%d');
my $missingMembershipMap = {};
my $row = 1;
process_data_file(
  'data/member_orders.csv',
  sub {
    my $values = shift;
    # dd($values);

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
    $values->{'CampaignCode'} = '';
    $values->{'FundCode'} = '';
    $values->{'AppealCode'} = '';
    $values->{'CommentsOnInvoice'} = '';
    $values->{'Comments'} = '';

    $values->{'InvoiceDescription'} = $values->{'MembershipTypeDes'};

    $values->{'ShipCustomerId'} = $values->{'PerBillableMemberId'};

    my $membershipTypeKey = uc $values->{'MembershipTypeDes'};
    my $paymentMethodKey = uc $values->{'PaymentMethod'};

    my $taxRate = $taxRates->{'Other'};

    if (exists($taxRates->{$values->{'MembershipBranch'}})) {
      $taxRate = $taxRates->{$values->{'MembershipBranch'}};
    }

    $values->{'FullfillStatusDate'} = $values->{'OrderDate'};

    $values->{'TrxInvoiceId'} = '';
    $values->{'ProductCode'} = '';
    $values->{'RateCode'} = '';
    $values->{'MarketCode'} = '';
    $values->{'DiscountAmount'} = 0;
    $values->{'DiscountCode'} = '';

    my $discount = 0;
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
    
    my $baseFee = $values->{'RenewMembershipFee'} || 0;

    if ($membershipTypeKey =~ /PRD/) {
      (my $prd = $membershipTypeKey) =~ s/\-.*//;
      $values->{'RateCode'} = 'Monthly';
      $values->{'PaymentMethod'} = 'Monthly';
      unless (exists($prdRates->{$prd}{$values->{'RateCode'}})) {
        print "Missing PRD mapping for '$prd' $values->{'RateCode'}\n";
        dd($values);
      }
      $baseFee = $prdRates->{$prd}{$values->{'RateCode'}};
    }

    $missingMembershipMap->{$membershipTypeKey}{$paymentMethodKey}++ unless ($values->{'RateCode'});

    my $orderDate = ParseDate($values->{'OrderDate'});
    $values->{'BeginDate'} = UnixDate($orderDate, '%Y-%m-%d');
    my $cycle = $cycleDurations->{$values->{'PaymentMethod'}};    
    $values->{'EndDate'} = UnixDate(DateCalc(DateCalc($orderDate, '+' . $cycle), '-1 day'), '%Y-%m-%d');

    $values->{'DiscountAmount'} = 0;
    if ($discount =~ /\%/) {
      $discount =~ s/\%//;
      $discount /= 100;

      $values->{'DiscountAmount'} = $baseFee * $discount;
    } elsif ($discount =~ /[\$,]/) {
      $discount =~ s/[\$,]//g;

      $values->{'DiscountAmount'} = $discount;
    }

    my $finalFee = $baseFee - $values->{'DiscountAmount'};
    $finalFee = 0 if ($finalFee < 0);

    $values->{'TaxPaidAmount'} = sprintf("%.2f", $finalFee * $taxRate);

    $values->{'TotalAmount'} = $finalFee + $values->{'TaxPaidAmount'};
    $values->{'PaymentAmount'} = $values->{'TotalAmount'};

    if ($values->{'AccessDenied'} eq 'Deny') {
      $values->{'ProductCode'} = 'AS_GMV_SO_SO';
      $values->{'EndDate'} = $denyEndDate;
    }

    check_order_errors('Memberships', $values);
    
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
    $values->{'CampaignCode'} = '';
    $values->{'FundCode'} = '';
    $values->{'AppealCode'} = '';
    $values->{'CommentsOnInvoice'} = '';
    $values->{'Comments'} = '';
    $values->{'FullfillStatusDate'} = '';
    $values->{'PerSolicitorId'} = '';
    
    $values->{'InvoiceDescription'} = $values->{'ItemDescription'};

    $values->{'ShipCustomerId'} = $values->{'PerMemberId'};

    $values->{'BeginDate'} = UnixDate($values->{'ProgramStartDate'}, '%Y-%m-%d');
    $values->{'EndDate'} = UnixDate($values->{'ProgramEndDate'}, '%Y-%m-%d');

    $values->{'TrxInvoiceId'} = $values->{'ReceiptNumber'};

    $values->{'TotalAmount'} = $values->{'FeePaid'};
    $values->{'PaymentAmount'} = $values->{'TotalAmount'};

    $values->{'ParentProductCode'} = $values->{'ProductCode'};

    check_order_errors('Programs', $values);
    
    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);

my @fndOrderAllColumns = get_template_columns($fndOrderTemplateName);

my $fndOrderWorkbook = make_workbook($fndOrderTemplateName);
my $fndOrderWorksheet = make_worksheet($fndOrderWorkbook, \@fndOrderAllColumns);
my $fndOrderRow = 1;

my @fndHardCreditAllColumns = get_template_columns($fndHardCreditTemplateName);

my $fndHardCreditWorkbook = make_workbook($fndHardCreditTemplateName);
my $fndHardCreditWorksheet = make_worksheet($fndHardCreditWorkbook, \@fndHardCreditAllColumns);
my $fndHardCreditRow = 1;

my $nextDueDate = UnixDate(DateCalc(ParseDate('today'), '+1 month'), '%Y-%m-%d');

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
    $values->{'FullfillStatusDate'} = '';
    $values->{'SoftCreditCustId'} = '';
    $values->{'TributeTypeCode'} = '';
    $values->{'InTributeToCustId'} = '';
    $values->{'InTributeToCustName'} = '';
    $values->{'AppealCode'} = '';

    $values->{'Anonymous'} = $values->{'DonorName'} =~ /ann?on/i ? 'Y' : 'N';

    $values->{'InvoiceDescription'} = $values->{'ItemDescription'};

    $values->{'Comments'} = decode_base64($values->{'Comments'});
    $values->{'CommentsOnInvoice'} = 'N';
    $values->{'CommentsOnInvoice'} = 'Y' if ($values->{'Comments'});

    $values->{'ShipCustomerId'} = $values->{'PerMemberId'};

    $values->{'BeginDate'} = '';
    $values->{'EndDate'} = '';

    $values->{'DueDate'} = $nextDueDate;

    $values->{'TrxInvoiceId'} = $values->{'ReceiptNumber'};

    $values->{'TotalAmount'} = $values->{'FeePaid'};
    $values->{'PaymentAmount'} = $values->{'FeePaid'} - $values->{'Balance'};

    $values->{'ParentProductCode'} = $values->{'ProductCode'};

    check_order_errors('Campaign', $values);

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );

    write_record(
      $fndOrderWorksheet,
      $fndOrderRow++,
      make_record($values, \@fndOrderAllColumns, $fndOrderColumnMap)
    );

    write_record(
      $fndHardCreditWorksheet,
      $fndHardCreditRow++,
      make_record($values, \@fndHardCreditAllColumns, $fndHardCreditColumnMap)
    );
  }
);

process_data_file(
  'data/arbal_orders.csv',
  sub {
    my $values = shift;
    # dd($values);

    $values->{'SubSystem'} = 'MISC';
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
    $values->{'BackIssueFlag'} = 'N';
    $values->{'PricingDiscountCode'} = 'USD';
    $values->{'FullfillStatusDate'} = '';
    $values->{'SoftCreditCustId'} = '';
    $values->{'TributeTypeCode'} = '';
    $values->{'InTributeToCustId'} = '';
    $values->{'InTributeToCustName'} = '';
    $values->{'CampaignCode'} = '';
    $values->{'FundCode'} = '';
    $values->{'AppealCode'} = '';

    $values->{'Anonymous'} = '';

    $values->{'InvoiceDescription'} = $values->{'ItemDescription'};

    $values->{'Comments'} = $values->{'Comments'};
    $values->{'CommentsOnInvoice'} = 'N';
    $values->{'CommentsOnInvoice'} = 'Y' if ($values->{'Comments'});

    $values->{'ShipCustomerId'} = $values->{'PerMemberId'};

    $values->{'BeginDate'} = '';
    $values->{'EndDate'} = '';

    $values->{'DueDate'} = '';

    $values->{'TrxInvoiceId'} = $values->{'ReceiptNumber'};

    $values->{'TotalAmount'} = $values->{'Balance'};
    $values->{'PaymentAmount'} = 0;

    $values->{'ParentProductCode'} = $values->{'ProductCode'};

    check_order_errors('Campaign', $values);

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );

  }
);

sub check_order_errors {
  my $source = shift;
  my $values = shift;

  our $orderProblemsWorksheet;
  our $orderProblemRow;

  unless ($values->{'InvoiceDescription'}) {
    write_record($orderProblemsWorksheet, $orderProblemRow++, [
      $source,
      $values->{'MemberId'} || '',
      $values->{'MembershipTypeDes'} || '',
      'Missing description',
    ]);
  }

  if (!exists($values->{'DiscountAmount'}) || $values->{'DiscountAmount'} eq '') {
    write_record($orderProblemsWorksheet, $orderProblemRow++, [
      $source,
      $values->{'MemberId'} || '',
      $values->{'MembershipTypeDes'} || '',
      'Null discount amount',
    ]);
  }

  if ($values->{'TotalAmount'} < 0) {
    write_record($orderProblemsWorksheet, $orderProblemRow++, [
      $source,
      $values->{'MemberId'} || '',
      $values->{'MembershipTypeDes'} || '',
      'Total amount less than zero',
    ]);
  }

  if ($values->{'TaxPaidAmount'} < 0) {
    write_record($orderProblemsWorksheet, $orderProblemRow++, [
      $source,
      $values->{'MemberId'} || '',
      $values->{'MembershipTypeDes'} || '',
      'Tax amount less than zero',
    ]);
  }

}