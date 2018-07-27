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
};

my $fndOrderTemplateName = 'DCT_FND_ORDER_DETAIL-67574';

my $fndOrderColumnMap = {
  'ORDER_NO'                               => { 'type' => 'record', 'source' => 'OrderNo' },
};

my $fndHardCreditTemplateName = 'DCT_FND_HARD_CREDIT-95863';

my $row = 1;
process_data_file(
  'data/member_orders.csv',
  sub {
    my $values = shift;
    # dd($values);

    $values->{'SubSystem'} = 'MBR';
    $values->{'ParentProductCode'} = 'GMV';
    $values->{'TaxableFlag'} = 'Y';
    $values->{'TaxCategoryCode'} = 'SALES_TAX';
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
      return if ($values->{'SponsorDiscount'} eq '0');
      $values->{'DiscountCode'} = 'SPONSOR' . $values->{'SponsorDiscount'};
      $discount = $values->{'SponsorDiscount'} . '%';
    }
    
    my $baseFee = $values->{'RenewMembershipFee'} || 0;
    my $isPrd = 0;

    if ($membershipTypeKey =~ /PRD/) {
      (my $prdMembership = $membershipTypeKey) =~ s/\-.*(PRD .)//;
      my $prdType = $1;

      $isPrd = 1;

      $values->{'RateCode'} = 'Monthly';
      $values->{'PaymentMethod'} = 'Monthly';
      
      unless (exists($prdRates->{$prdMembership}{$values->{'RateCode'}})) {
        print "Missing PRD mapping for '$prdMembership' $values->{'RateCode'}\n";
        dd($values);
      }
      unless (exists($prdDiscounts->{$prdType})) {
        print "Missing PRD discount for '$prdType'\n";
        dd($values);
      }

      $baseFee = $prdRates->{$prdMembership}{$values->{'RateCode'}};

      $values->{'DiscountAmount'} = $prdDiscounts->{$prdType};
      
      # print "PRD ($prdMembership) base fee is $baseFee (adjust for $prdType)\n";
      # dd($values);
      # exit;
    }

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

    # Include tax in total or not?
    $values->{'TotalAmount'} = $finalFee + $values->{'TaxPaidAmount'};
    # $values->{'TotalAmount'} = $finalFee;
    $values->{'PaymentAmount'} = $isPrd ? 0 : $values->{'TotalAmount'};

    dd($values) if ($values->{'PaymentAmount'} < 0);

    if ($values->{'AccessDenied'} eq 'Deny') {
      $values->{'ProductCode'} = 'AS_GMV_SO_SO';
      $values->{'RateCode'} = 'Annual';
      $values->{'EndDate'} = $denyEndDate;
    }

    $missingMembershipMap->{$membershipTypeKey}{$paymentMethodKey}++ unless ($values->{'RateCode'});

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

    $values->{'TotalAmount'} = $values->{'OrderTotal'};
    $values->{'PaymentAmount'} = $values->{'OrderTotal'} - $values->{'BalanceDue'};

    if ($values->{'PaymentAmount'} < 0) {
      print "Order $values->{'OrderNo'} for $values->{'PerMemberId'} ($values->{'ProductCode'}) was $values->{'PaymentAmount'}\n";
    }

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

    $values->{'TotalAmount'} = $values->{'OrderTotal'};
    $values->{'PaymentAmount'} = $values->{'OrderTotal'} - $values->{'BalanceDue'};

    dd($values) if ($values->{'PaymentAmount'} < 0);

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

    $values->{'TotalAmount'} = $values->{'BalanceDue'};
    $values->{'PaymentAmount'} = 0;

    dd($values) if ($values->{'PaymentAmount'} < 0);

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