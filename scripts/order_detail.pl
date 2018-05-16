#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Date::Manip;

my $templateName = 'DCT_ORDER_DETAIL-32358';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'orderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ORDER_DATE'                         => { 'type' => 'record', 'source' => 'orderDate' },
  'SHIP_CUSTOMER_ID'                   => { 'type' => 'record', 'source' => 'shipCustomerId' },
  'SHIP_ADDRESS_TYPE_CODE'             => { 'type' => 'static', 'source' => 'HOME' },
  'INVOICE_NO'                         => { 'type' => 'record', 'source' => 'trxInvoiceId' },
  'INVOICE_DATE'                       => { 'type' => 'record', 'source' => 'orderDate' },
  'SUBSYSTEM'                          => { 'type' => 'static', 'source' => 'MBR' },
  'PRODUCT_CODE'                       => { 'type' => 'record', 'source' => 'productCode' },
  'PARENT_PRODUCT'                     => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'LINE_TYPE'                          => { 'type' => 'static', 'source' => 'IP' },
  'LINE_STATUS_CODE'                   => { 'type' => 'static', 'source' => 'A' },
  'LINE_STATUS_DATE'                   => { 'type' => 'record', 'source' => 'orderDate' },
  'FULFILL_STATUS_CODE'                => { 'type' => 'static', 'source' => 'A' },
  'FULFILL_STATUS_DATE'                => { 'type' => 'record', 'source' => 'orderDate' },
  'RECOGNITION_STATUS_CODE'            => { 'type' => 'static', 'source' => 'C' },
  'RATE_STRUCTURE'                     => { 'type' => 'static', 'source' => 'LIST' },
  'RATE_CODE'                          => { 'type' => 'record', 'source' => 'rateCode' },
  'TAXABLE_FLAG'                       => { 'type' => 'static', 'source' => 'Y' },
  'TAX_CATEGORY_CODE'                  => { 'type' => 'static', 'source' => 'SALES' },
  'REQUESTED_QTY'                      => { 'type' => 'static', 'source' => '1' },
  'ORDER_QTY'                          => { 'type' => 'static', 'source' => '1' },
  'TOTAL_AMOUNT'                       => { 'type' => 'record', 'source' => 'totalAmount' },
  'CYCLE_BEGIN_DATE'                   => { 'type' => 'record', 'source' => 'beginDate' },
  'CYCLE_END_DATE'                     => { 'type' => 'record', 'source' => 'endDate' },
  'BACK_ISSUE_FLAG'                    => { 'type' => 'static', 'source' => 'Y' },
  'INITIAL_BEGIN_DATE'                 => { 'type' => 'record', 'source' => 'joinDate' },
  'RETURNED_QTY'                       => { 'type' => 'static', 'source' => '0' },
  'PAYOR_CUSTOMER_ID'                  => { 'type' => 'record', 'source' => 'billCustomerId' },
  'RECEIPT_TYPE'                       => { 'type' => 'static', 'source' => 'CASH' },
  'RECEIPT_CURRENCY_CODE'              => { 'type' => 'static', 'source' => 'USD' },
  'RECEIPT_DATE'                       => { 'type' => 'record', 'source' => 'orderDate' },
  'XRATE'                              => { 'type' => 'static', 'source' => '1' },
  'PAYMENT_AMOUNT'                     => { 'type' => 'record', 'source' => 'totalAmount' },
  'RECEIPT_STATUS_CODE'                => { 'type' => 'static', 'source' => 'A' },
  'RECEIPT_STATUS_DATE'                => { 'type' => 'record', 'source' => 'orderDate' },
  'CL_LATE_FEE_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'MANUAL_DISCOUNT_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'DISCOUNT_CODE'                      => { 'type' => 'record', 'source' => 'discountCode' },
  'ACTUAL_DISCOUNT_AMOUNT'             => { 'type' => 'record', 'source' => 'discountAmount' },
  'ACTUAL_SHIP_AMOUNT'                 => { 'type' => 'static', 'source' => '0' },
  'ACTUAL_TAX_AMOUNT'                  => { 'type' => 'record', 'source' => 'taxPaidAmount' },
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

my @orderMasterFields = qw(
  orderNo
  orderDate
  orgId
  orgUnitId
  billCustomerId
  billAddressTypeCode
  shipCustomerId
  orderMethodCode
  orderStatusCode
  orderStatusDate
  ordstsReasonCode
  clOrderMethodCode
  couponCode
  application
  ackLetterMethodCode
  poNumber
  confirmationNo
  ackLetterPrintDate
  confirmationDate
  orderCompleteFlag
  advContractId
  advAgencyCustId
  billSalesTerritory
  fndGiveEmployerCreditFlag
  shipSalesTerritory
  posFlag
  posCountryCode
  posState
  posPostalCode
  advRateCardYearCode
  advAgencySubCustId
  employerCustomerId
  oldOrderNo
  nextBillDate
  joinDate
);

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

open(my $orderMaster, '<', 'data/order_master.txt')
  or die "Couldn't open data/order_master.txt: $!";
<$orderMaster>; # eat the headers

my $row = 1;
while(<$orderMaster>) {
  chomp;
  my $values = split_values($_, @orderMasterFields);

  my $nextBillDate = ParseDate($values->{'nextBillDate'});
  $values->{'beginDate'} = UnixDate(DateCalc($nextBillDate, '+1 day'), '%Y-%m-%d');
  $values->{'endDate'} = UnixDate(DateCalc($nextBillDate, '-1 day'), '%Y-%m-%d');

  $values->{'trxInvoiceId'} = '';
  $values->{'productCode'} = '';
  $values->{'rateCode'} = '';
  $values->{'totalAmount'} = '';
  $values->{'discountAmount'} = '';
  $values->{'taxPaidAmount'} = '';

  write_record(
    $worksheet,
    $row++,
    make_record($values, \@allColumns, $columnMap)
  );

}

close($orderMaster);
