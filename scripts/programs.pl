#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::CSV_XS;
use Date::Manip;
use Term::ProgressBar;

my $productTemplateName = 'DCT_MTG_PRODUCT-73559';

my $productColumnMap = {
  'PRODUCT_CODE'                             => { 'type' => 'record', 'source' => 'ProductCode' },
  'ORG_ID'                                   => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                              => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'MEETING_START_DATE'                       => { 'type' => 'record', 'source' => 'StartDateTime' },
  'MEETING_END_DATE'                         => { 'type' => 'record', 'source' => 'EndDateTime' },
  'LAST_REGISTRATION_DATE'                   => { 'type' => 'record', 'source' => 'LastRegistrationDate' },
  'LAST_REFUND_DATE'                         => { 'type' => 'record', 'source' => 'LastRefundDate' },
  'CAPACITY'                                 => { 'type' => 'record', 'source' => 'MaxCapacity' },
  'REGISTRATIONS'                            => { 'type' => 'static', 'source' => '0' },
  'WAIT_LIST_CAPACITY'                       => { 'type' => 'static', 'source' => '99' },
  'WAIT_LIST_REGISTRATIONS'                  => { 'type' => 'static', 'source' => '0' },
  'ALLOW_SESSION_CONFLICT_FLAG'              => { 'type' => 'static', 'source' => 'N' },
  'ENFORCE_MASTER_BLOCK_FLAG'                => { 'type' => 'static', 'source' => 'N' },
  'ENFORCE_INVENTORY_FLAG'                   => { 'type' => 'static', 'source' => 'N' },
  'ROOM_NUMBER'                              => { 'type' => 'static', 'source' => '0' },
  'INVOICE_DESCRIPTION'                      => { 'type' => 'record', 'source' => 'ProgramDescription' },
  'LONG_NAME'                                => { 'type' => 'record', 'source' => 'ProgramDescription' },
  'PRODUCT_TYPE_CODE'                        => { 'type' => 'static', 'source' => 'M' },
  'PRODUCT_CLASS_CODE'                       => { 'type' => 'record', 'source' => 'Department' },
  'PRODUCT_STATUS_CODE'                      => { 'type' => 'static', 'source' => 'A' },
  'PRODUCT_STATUS_DATE'                      => { 'type' => 'record', 'source' => 'AvailableDate' },
  'AVAILABLE_DATE'                           => { 'type' => 'record', 'source' => 'AvailableDate' },
  'EXPIRATION_DATE'                          => { 'type' => 'record', 'source' => 'ExpirationDate' },
  'MASTER_PRODUCT_FLAG'                      => { 'type' => 'static', 'source' => 'Y' },
  'AVAILABLE_TO_ORDERS_FLAG'                 => { 'type' => 'static', 'source' => 'Y' },
  'TAXABLE_FLAG'                             => { 'type' => 'static', 'source' => 'N' },
  'REVENUE_RECOG_METHOD_CODE'                => { 'type' => 'static', 'source' => 'BEGIN' },
  'PAY_PRIORITY'                             => { 'type' => 'static', 'source' => '0' },
  'AR_ACCOUNT'                               => { 'type' => 'record', 'source' => 'ArAccount' },
  'PPL_ACCOUNT'                              => { 'type' => 'record', 'source' => 'PplAccount' },
  'WRITEOFF_ACCOUNT'                         => { 'type' => 'record', 'source' => 'WriteOffAccount' },
  'CANCELLATION_ACCOUNT'                     => { 'type' => 'record', 'source' => 'CancelAccount' },
  'DISCOUNT_ACCOUNT'                         => { 'type' => 'record', 'source' => 'DiscountAccount' },
  'DEFERRED_DISCOUNT_ACCOUNT'                => { 'type' => 'record', 'source' => 'DefDiscountAccount' },
  'AGENCY_DISC_ACCOUNT'                      => { 'type' => 'record', 'source' => 'AgencyDiscountAccount' },
  'AGENCY_DEFERRED_DISC_ACCOUNT'             => { 'type' => 'record', 'source' => 'AgencyDefDiscountAccount' },
  'RATE_STRUCTURE'                           => { 'type' => 'static', 'source' => 'LIST' },
  'RATE_CODE'                                => { 'type' => 'static', 'source' => 'STD' },
  'RATE_CODE_DESCR'                          => { 'type' => 'static', 'source' => 'Standard' },
  'RATE_CODE_DISPLAY_ORDER'                  => { 'type' => 'static', 'source' => '0' },
  'SHORT_PAY_PROC_CODE'                      => { 'type' => 'record', 'source' => 'ShortPayProcCode' },
  'AGENCY_DISCOUNT_PCT'                      => { 'type' => 'static', 'source' => '0' },
  'ECOMMERCE_FLAG'                           => { 'type' => 'static', 'source' => 'Y' },
  'DEFAULT_RATE_WEB_FLAG'                    => { 'type' => 'static', 'source' => 'Y' },
  'WAIVE_SHIPPING_FLAG'                      => { 'type' => 'static', 'source' => 'Y' },
  'PRICE_CURRENCY_CODE'                      => { 'type' => 'static', 'source' => 'USD' },
  'PRICE_BEGIN_DATE'                         => { 'type' => 'record', 'source' => 'AvailableDate' },
  'PRICE'                                    => { 'type' => 'record', 'source' => 'ListPrice' },
  'CL_VARIABLE_PART_SCHEDULE_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'MIN_PRICE'                                => { 'type' => 'static', 'source' => '0' },
  'MAX_PRICE'                                => { 'type' => 'static', 'source' => '0' },
  'CL_MIN_SESSIONS'                          => { 'type' => 'static', 'source' => '0' },
  'CL_TARGET_REGISTRATIONS'                  => { 'type' => 'static', 'source' => '0' },
  'WRITEOFF_TOLERANCE'                       => { 'type' => 'static', 'source' => '0' },
  'TAXABLE_AMOUNT'                           => { 'type' => 'static', 'source' => '0' },
  'CL_TARGET_BUDGET'                         => { 'type' => 'static', 'source' => '0' },
  'REVENUE_ACCTS_EFFECTIVE_BEGIN_DATE'       => { 'type' => 'record', 'source' => 'AvailableDate' },
  'REVENUE_ACCOUNT_1'                        => { 'type' => 'record', 'source' => 'RevenueAccount' },
  'DEFERRED_ACCOUNT_1'                       => { 'type' => 'record', 'source' => 'DeferredAccount' },
  'DISTRIBUTION_PCT_1'                       => { 'type' => 'static', 'source' => '100' },
  'CL_CCC_FLAG'                              => { 'type' => 'record', 'source' => 'ClCccFlag' },
  'CL_BRANCH'                                => { 'type' => 'record', 'source' => 'BranchCode' },
  'CL_MON'                                   => { 'type' => 'record', 'source' => 'ClMon' },
  'CL_TUE'                                   => { 'type' => 'record', 'source' => 'ClTue' },
  'CL_WED'                                   => { 'type' => 'record', 'source' => 'ClWed' },
  'CL_THU'                                   => { 'type' => 'record', 'source' => 'ClThu' },
  'CL_FRI'                                   => { 'type' => 'record', 'source' => 'ClFri' },
  'CL_SAT'                                   => { 'type' => 'record', 'source' => 'ClSat' },
  'CL_SUN'                                   => { 'type' => 'record', 'source' => 'ClSun' },
  'CL_ENFORCE_AGE'                           => { 'type' => 'static', 'source' => 'N' },
  'CL_ENFORCE_GENDER'                        => { 'type' => 'static', 'source' => 'N' },
  'CL_ENFORCE_GRADE'                         => { 'type' => 'static', 'source' => 'N' },
  'DIRECT_PRICE_UPDATE_FLAG'                 => { 'type' => 'static', 'source' => 'N' },
  'MEMBERS_ONLY_FLAG'                        => { 'type' => 'static', 'source' => 'N' },
  'CL_MIN_MON'                               => { 'type' => 'record', 'source' => 'MinAgeMonths' },
  'CL_MAX_MON'                               => { 'type' => 'record', 'source' => 'MaxAgeMonths' },
  'AR_ACCOUNT_COMPANY_NUMBER'                => { 'type' => 'static', 'source' => '1' },
  'CL_MIN_YRS'                               => { 'type' => 'record', 'source' => 'MinAgeYears' },
  'CL_MAX_YRS'                               => { 'type' => 'record', 'source' => 'MaxAgeYears' },
  'INCLUDE_IN_PROGRAM_FLAG'                  => { 'type' => 'static', 'source' => 'Y' },
  'MULTI_DAY_FLAG'                           => { 'type' => 'static', 'source' => 'Y' },
  'ALLOW_CAPACITY_OVERRIDE_FLAG'             => { 'type' => 'static', 'source' => 'N' },
  'CL_ANS_REQUIRED1'                         => { 'type' => 'static', 'source' => 'N' },
  'DONATION_FLAG'                            => { 'type' => 'static', 'source' => 'N' },
  'HAS_ASSIGNED_SALES_REP_FLAG'              => { 'type' => 'static', 'source' => 'N' },
  'CL_ANS_REQUIRED2'                         => { 'type' => 'static', 'source' => 'N' },
  'CL_DEPARTMENT'                            => { 'type' => 'record', 'source' => 'Department' },
  'CL_DEPARTMENT_SUBCLASS'                   => { 'type' => 'record', 'source' => 'DepartmentSubClass' },
  'ATTENDANCE_DEFAULT_FLAG'                  => { 'type' => 'static', 'source' => 'Y' },
  'TICKETED_EVENT_FLAG'                      => { 'type' => 'static', 'source' => 'N' },
  'ZEROPRICE_FLAG'                           => { 'type' => 'static', 'source' => 'N' },
  'EXCLUDE_FROM_DISCOUNT_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'TIME_ZONE_CODE'                           => { 'type' => 'static', 'source' => 'UNDEFINED' },
  'LIMITED_SEATS_THRESHOLD'                  => { 'type' => 'record', 'source' => 'LimitedSeatsThreshold' },
  'DISPLAY_EMERGENCY_CONTACT_INFO_FLAG'      => { 'type' => 'static', 'source' => 'Y' },
  'DISPLAY_SPECIAL_NEEDS_CONTROL_FLAG'       => { 'type' => 'static', 'source' => 'Y' },
  'CL_PRODUCT_SUBCLASS_CODE'                 => { 'type' => 'record', 'source' => 'DepartmentSubClass' },
  'CREATE_SESSION_DETAIL_PAGE_FLAG'          => { 'type' => 'static', 'source' => 'N' },
  'DAILY_FLAG'                               => { 'type' => 'static', 'source' => 'N' },
  'AVAILABLE_FOR_ALL_DAILY_RATES_FLAG'       => { 'type' => 'static', 'source' => 'N' },
  'WEB_DISPLAY_REGISTRANT_CONTACT_INFO_FLAG' => { 'type' => 'static', 'source' => 'Y' },
  'ALLOW_REGISTRATION_OF_OTHERS_FLAG'        => { 'type' => 'static', 'source' => 'N' },
};

my @productAllColumns = get_template_columns($productTemplateName);

my $productWorkbook = make_workbook($productTemplateName);
my $productWorksheet = make_worksheet($productWorkbook, \@productAllColumns);

my $addRateCodeTemplateName = 'DCT_MTG_ADDITIONAL_RATE_CODE-83797';

my $addRateCodeColumnMap = {
  'PRODUCT_CODE'                  => { 'type' => 'record', 'source' => 'ProductCode' },
  'PARENT_PRODUCT'                => { 'type' => 'record', 'source' => 'ProductCode' },
  'ORG_ID'                        => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                   => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'RATE_STRUCTURE'                => { 'type' => 'static', 'source' => 'MBR' },
  'RATE_CODE'                     => { 'type' => 'static', 'source' => 'STD' },
  'WAIVE_SHIPPING_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'SHORT_PAY_PROC_CODE'           => { 'type' => 'record', 'source' => 'ShortPayProcCode' },
  'AGENCY_DISCOUNT_PCT'           => { 'type' => 'static', 'source' => '0' },
  'DEFAULT_RATE_FLAG'             => { 'type' => 'static', 'source' => 'Y' },
  'SORT_ORDER'                    => { 'type' => 'static', 'source' => '0' },
  'DEFAULT_RATE_WEB_FLAG'         => { 'type' => 'static', 'source' => 'Y' },
  'ECOMMERCE_FLAG'                => { 'type' => 'static', 'source' => 'Y' },
  'ACTIVE_FLAG'                   => { 'type' => 'static', 'source' => 'Y' },
  'RATE_CODE_DESCR'               => { 'type' => 'static', 'source' => 'Standard' },
  'PRORATE_AMOUNT_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'BACK_ISSUES_FLAG'              => { 'type' => 'static', 'source' => 'N' },
};

my @addRateCodeAllColumns = get_template_columns($addRateCodeTemplateName);

my $addRateCodeWorkbook = make_workbook($addRateCodeTemplateName);
my $addRateCodeWorksheet = make_worksheet($addRateCodeWorkbook, \@addRateCodeAllColumns);

my $addPriceTemplateName = 'DCT_MTG_ADDITIONAL_PRICING-46120';

my $addPriceColumnMap = {
  'PRODUCT_CODE'       => { 'type' => 'record', 'source' => 'ProductCode' },
  'PARENT_PRODUCT'     => { 'type' => 'record', 'source' => 'ProductCode' },
  'ORG_ID'             => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'        => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'RATE_STRUCTURE'     => { 'type' => 'static', 'source' => 'MBR' },
  'RATE_CODE'          => { 'type' => 'static', 'source' => 'STD' },
  'CURRENCY_CODE'      => { 'type' => 'static', 'source' => 'USD' },
  'PRICE_BEGIN_DATE'   => { 'type' => 'record', 'source' => 'AvailableDate' },
  'PRICE'              => { 'type' => 'record', 'source' => 'FullMemberPrice' },
  'MIN_PRICE'          => { 'type' => 'static', 'source' => '0' },
  'MAX_PRICE'          => { 'type' => 'static', 'source' => '0' },
  'WRITEOFF_TOLERANCE' => { 'type' => 'static', 'source' => '0' },
  'TAXABLE_AMOUNT'     => { 'type' => 'static', 'source' => '0' },
  'FAIR_MARKET_VALUE'  => { 'type' => 'record', 'source' => 'FullMemberPrice' },
};

my @addPriceAllColumns = get_template_columns($addPriceTemplateName);

my $addPriceWorkbook = make_workbook($addPriceTemplateName);
my $addPriceWorksheet = make_worksheet($addPriceWorkbook, \@addPriceAllColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

my($dataFile, $headers, $totalRows) = open_data_file('data/ProductCodes.csv', programCodesHeaderMap());

our $programCodes = {};
while(my $rowIn = $csv->getline($dataFile)) {

  my $values = map_values($headers, $rowIn);
  # dump($values); exit;
  die "Duplicate program subdepartment: $values->{'SubDepartmentName'}"
    if (exists($programCodes->{$values->{'SubDepartmentName'}}));

  $values->{'SubDepartmentName'} =~ s/^\s+|\s+$//g;
  $programCodes->{lc $values->{'SubDepartmentName'}} = $values;
}

close($dataFile);

my $products = [];

($dataFile, $headers, $totalRows) = open_data_file('data/Programs.csv', programsHeaderMap());

print "Processing programs\n";
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });

my $count = 1;
while(my $rowIn = $csv->getline($dataFile)) {

  $progress->update($count++);

  my $values = clean_program_values(map_values($headers, $rowIn));
  # dump($values); exit;
  
  next if (skip_cycle($values->{'ProgramType'}));
  next unless ($values->{'Active'} eq 'YES');

  push(@{$products}, $values);
}

close($dataFile);

($dataFile, $headers, $totalRows) = open_data_file('data/Camp.csv', campHeaderMap());

print "Processing camps\n";
$progress = Term::ProgressBar->new({ 'count' => $totalRows });

$count = 1;
while(my $rowIn = $csv->getline($dataFile)) {

  $progress->update($count++);

  my $values = clean_camp_values(map_values($headers, $rowIn));
  # dump($values); exit;

  push(@{$products}, $values);
}

close($dataFile);

my $programTypeWorkbook = make_workbook('unmatched_program_type');
my $programTypeWorksheet = make_worksheet($programTypeWorkbook, 
  ['Source', 'Program No', 'Type', 'Description', 'Session Start Date', 'Schedule']);

my $noStartDateWorkbook = make_workbook('missing_start_date');
my $noStartDateWorksheet = make_worksheet($noStartDateWorkbook, 
  ['Start', 'End', 'Program', 'Description', 'Summary', 'Session Start Date', 
    'Session End Date', 'Class Start Time', 'Duration', 'Week Days']);

my $startEndDatesWorkbook = make_workbook('start_end_dates');
my $startEndDatesWorksheet = make_worksheet($startEndDatesWorkbook, 
  ['Product Code', 'Branch', 'Program', 'Description', 'Summary', 'Session Start Date', 
    'Session End Date', 'Class Start Time', 'Duration', 'Week Days', 'Start', 'End']);

my $productCodes = {};
my $productsByCode = {};

my $collector = {};
print "Generating program files\n";
our $partTracker = {};
my $availableDate = UnixDate(ParseDate('1/1/2000'), '%Y-%m-%d');

my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');
$dbh->do(q{
  delete from products
  });

$progress = Term::ProgressBar->new({ 'count' => scalar(@{$products}) });
$count = 1;
my $row = 1;
my $programTypeRow = 1;
my $startEndDatesRow = 1;
my $missingStartRow = 1;
foreach my $program (sort { by_program_start_date($a, $b) } @{$products}) {
  $progress->update($count++);

  $program->{'AvailableDate'} = $availableDate;
  
  my $productDetails = get_product_details($program);

  $program->{'ProductCode'} = $productDetails->{'ProductCode'};
  $program->{'Department'} = $productDetails->{'Department'};
  $program->{'DepartmentSubClass'} = $productDetails->{'DepartmentSubClass'};

  unless ($program->{'ProductCode'}) {
    write_record($programTypeWorksheet, $programTypeRow++, [
      $program->{'Source'},
      $program->{'ProgramNumber'} || '',
      $program->{'ProgramType'} || '',
      $program->{'ProgramDescription'} || '',
      $program->{'SessionStartDate'} || '',
      $program->{'Schedule'} || '',
    ]);

    next;
  }

  next if (skip_program($program->{'ProgramDescription'}));
  
  write_record($startEndDatesWorksheet, $startEndDatesRow++, [
    $program->{'ProductCode'} || '',
    $program->{'BranchName'} || '',
    $program->{'ProgramType'} || '',
    $program->{'ProgramDescription'} || '',
    $program->{'ItemDescription'} || $program->{'Session'} || '',
    $program->{'SessionStartDate'} || '',
    $program->{'SessionEndDate'} || '',
    $program->{'ClassStartTime'} || '',
    $program->{'ClassDuration'} || '',
    $program->{'WeekDays'} || '',
    $program->{'StartDateTime'} ? UnixDate($program->{'StartDateTime'}, '%a, %d %b %Y %i:%M:%S %p') : '',
    $program->{'EndDateTime'} ? UnixDate($program->{'EndDateTime'}, '%a, %d %b %Y %i:%M:%S %p') : '',
  ]);

  unless ($program->{'StartDateTime'} && $program->{'EndDateTime'}) {        
    write_record($noStartDateWorksheet, $missingStartRow++, [
      $program->{'StartDateTime'} || '',
      $program->{'EndDateTime'} || '',
      $program->{'ProgramType'} || '',
      $program->{'ProgramDescription'} || '',
      $program->{'ItemDescription'} || '',
      $program->{'SessionStartDate'} || '',
      $program->{'SessionEndDate'} || '',
      $program->{'ClassStartTime'} || '',
      $program->{'ClassDuration'} || '',
      $program->{'WeekDays'} || '',
    ]);

    next;
  }
  
  $productCodes->{$program->{'ProductCode'}}++;
  $productsByCode->{$program->{'ProductCode'}} = []
    unless (exists($productsByCode->{$program->{'ProductCode'}}));
  push(@{$productsByCode->{$program->{'ProductCode'}}}, $program);

  my $description = $program->{'ProgramDescription'};
  my $summary = $program->{'Summary'};

  $dbh->do(q{
    insert into products (product_code, branch, type, description, summary, session)
      values (?, ?, ?, ?, ?, ?)
    }, undef, $program->{'ProductCode'}, $program->{'BranchName'}, $program->{'Source'},
      $description, $summary, $program->{'Session'});

  my $productRecord = make_record($program, \@productAllColumns, $productColumnMap);
  write_record($productWorksheet, $row, $productRecord);

  my $addRateCodeRecord = make_record($program, \@addRateCodeAllColumns, $addRateCodeColumnMap);
  write_record($addRateCodeWorksheet, $row, $addRateCodeRecord);

  my $addPriceRecord = make_record($program, \@addPriceAllColumns, $addPriceColumnMap);
  write_record($addPriceWorksheet, $row, $addPriceRecord);

  $row++;
}

dump($collector) if (keys %{$collector});

my $duplicateProductCodeWorkbook = make_workbook('duplicate_product_code');
my $duplicateProductCodeWorksheet = make_worksheet($duplicateProductCodeWorkbook, 
  ['Product Code', 'Branch', 'Program', 'Description', 'Summary', 'Session Start Date', 
    'Session End Date', 'Class Start Time', 'Duration', 'Week Days', 'Start', 'End']);
my $duplicateProductCodeRow = 1;

foreach my $productCode (keys %{$productCodes}) {
  next unless ($productCodes->{$productCode} > 1);

  foreach my $program (@{$productsByCode->{$productCode}}) {
    write_record($duplicateProductCodeWorksheet, $duplicateProductCodeRow++, [
      $program->{'ProductCode'} || '',
      $program->{'BranchName'} || '',
      $program->{'ProgramType'} || '',
      $program->{'ProgramDescription'} || '',
      $program->{'ItemDescription'} || $program->{'Session'} || '',
      $program->{'SessionStartDate'} || '',
      $program->{'SessionEndDate'} || '',
      $program->{'ClassStartTime'} || '',
      $program->{'ClassDuration'} || '',
      $program->{'WeekDays'} || '',
      $program->{'StartDateTime'} || '',
      $program->{'EndDateTime'} || '',
    ]);
  }
}

sub get_product_details {
  my $program = shift;
  
  our $programCodes;

  my $productDetails = {
    'ProductCode' => '',
    'Department' => '',
    'DepartmentSubClass' => '',
    'Season' => '',
  };

  return $productDetails unless (
      exists($programCodes->{lc $program->{'MappedProgramDescription'}}) &&
      $program->{'SessionStartDate'}
    );

  my $codeInfo = $programCodes->{lc $program->{'MappedProgramDescription'}};
  
  $productDetails->{'Department'} = $codeInfo->{'ProductClass'};
  $productDetails->{'DepartmentSubClass'} = $codeInfo->{'SubClass'};
  $productDetails->{'Season'} = $codeInfo->{'Season'};

  if ($productDetails->{'Season'} eq 'M18' && $program->{'ProgramType'} =~ /Summer 2/i) {
    $productDetails->{'Season'} = 'M28';
  }

  my $increment = get_program_increment(
    $program->{'BranchCode'}, 
    $productDetails->{'Department'},
    $productDetails->{'DepartmentSubClass'},
    $program->{'Session'}
  );

  my @codeParts;
  push(@codeParts, $program->{'BranchCode'});
  push(@codeParts, $productDetails->{'Department'});
  push(@codeParts, $productDetails->{'DepartmentSubClass'});
  push(@codeParts, sprintf('%02s', $increment));
  push(@codeParts, UnixDate($program->{'SessionStartDate'}, '%m%d%y'));
  push(@codeParts, $productDetails->{'Season'});
  
  $productDetails->{'ProductCode'} = join('_', @codeParts);

  $productDetails->{'DepartmentSubClass'} .= '_';

  return $productDetails;
}

sub get_program_increment {
  my $branchCode = shift;
  my $department = shift || 'un';
  my $departmentSubCLass = shift || 'un';
  my $session = shift || '';

  our $partTracker;

  if ($session =~ /week (\d*)/i) {
    return $1;
  }

  return ++$partTracker->{$branchCode}{$department}{$departmentSubCLass};
}

sub clean_program_values {
  my $values = shift;

  $values->{'Source'} = 'program';

  $values->{'ClCccFlag'} = 'N';

  $values->{'SubsidyGl'} = '' if ($values->{'ScholarshipGlAccount'} eq 'NA');

  $values->{'Summary'} = $values->{'ItemDescription'};

  $values->{'ListPrice'} = $values->{'NonMemberPrice'};

  $values->{'SessionStartDate'} =~ s/ .*$//;
  $values->{'SessionEndDate'} =~ s/ .*$//;

  if ($values->{'ProgramDescription'} eq 'Little League & RBI Program'){
    $values->{'WeekDays'} = 'Saturday' if ($values->{'WeekDays'} =~ /^var/i);
    $values->{'ClassDuration'} = '1 hour' if ($values->{'ClassDuration'} =~ /^var/i);    
  }

  # Date::Manip doesn't seem to like high week durations.
  $values->{'ClassDuration'} = '140 days' if ($values->{'ClassDuration'} eq '20 weeks');
  $values->{'ClassDuration'} = '30 minutes' if ($values->{'ClassDuration'} eq '1/2 Hour');
  $values->{'ClassDuration'} =~ s/(hoiur|houtd|hr\.|hrs\.|hrs per day)/hours/;
  $values->{'ClassDuration'} =~ s/ 2x.*//i;
  $values->{'ClassDuration'} =~ s/\dk/4 hours/i;
  $values->{'ClassDuration'} =~ s/overnight/12 hours/i;
  $values->{'ClassDuration'} =~ s/mimutes/minutes/i;
  $values->{'ClassDuration'} =~ s/9:30.*-8.*/11 hours/i;
  $values->{'ClassDuration'} =~ s/6pm-8am/14 hours/i;
  $values->{'ClassDuration'} =~ s/^(\d+)$/$1 hours/i;
  $values->{'ClassDuration'} =~ s/^hour$/1 hour/i;

  my $startDate = get_start_date($values);
  my $endDate = get_end_date($values);

  $values->{'WeekDays'} =~ s/(TBD|TBA|On your own|Schedule Weigh-in|Vaires|Varies)//i;
  $values->{'WeekDays'} =~ s/M-R/MON,TUE,WED,THU/i;
  $values->{'WeekDays'} =~ s/M-W-F/MON,WED,FRI/i;
  $values->{'WeekDays'} =~ s/T-W-TH/TUE,WED,THU/i;
  $values->{'WeekDays'} =~ s/(Every ?day|Daily|all days)/SUN,MON,TUE,WED,THU,FRI,SAT/i;
  $values->{'WeekDays'} =~ s/(\/| & )/,/g;

  if ($values->{'WeekDays'} && $values->{'WeekDays'} !~ /^(mon|tue|wed|thu|fri|sat|sun)/i &&
      !(skip_program($values->{'ProgramDescription'}) || skip_cycle($values->{'ProgramType'}))) {
    print "Can't translate $values->{'WeekDays'}\n";
    $values->{'WeekDays'} = '';
  }
  my @daysOfWeek = split(',', lc $values->{'WeekDays'});

  my $startDayOfWeek = '';
  my $endDayOfWeek = '';
  if (@daysOfWeek) {
    $startDayOfWeek = $daysOfWeek[0];
    $endDayOfWeek = $daysOfWeek[$#daysOfWeek];
  }

  if ($startDate && $startDayOfWeek) {
    $values->{'StartDateTime'} = UnixDate(
      Date_GetNext(
        $startDate . ' ' . $values->{'ClassStartTime'}, 
        $startDayOfWeek, 
        1
      ), '%Y-%m-%d %r');
  }

  if ($endDate && $endDayOfWeek) {
    my $lastDayOfWeek = UnixDate(
      Date_GetPrev($endDate . ' ' . $values->{'ClassStartTime'}, $endDayOfWeek, 1), '%Y-%m-%d %r');

    $values->{'EndDateTime'} = UnixDate(DateCalc($lastDayOfWeek, 
      '+' . $values->{'ClassDuration'}), '%Y-%m-%d %r');
  }

  $values->{'ClMon'} = 'Y' if ($values->{'WeekDays'} =~ /Mon/i);
  $values->{'ClTue'} = 'Y' if ($values->{'WeekDays'} =~ /Tue/i);
  $values->{'ClWed'} = 'Y' if ($values->{'WeekDays'} =~ /Wed/i);
  $values->{'ClThu'} = 'Y' if ($values->{'WeekDays'} =~ /Thu/i);
  $values->{'ClFri'} = 'Y' if ($values->{'WeekDays'} =~ /Fri/i);
  $values->{'ClSat'} = 'Y' if ($values->{'WeekDays'} =~ /Sat/i);
  $values->{'ClSun'} = 'Y' if ($values->{'WeekDays'} =~ /Sun/i);

  $values->{'MappedProgramDescription'} = map_program_descriptions($values);

  return clean_all_values($values);
}

sub get_start_date {
  my $values = shift;

  my $startDate = $values->{'SessionStartDate'};

  return $startDate unless ($startDate eq '1/1/2018');

  # Look for the first "month" word in the summary and use it
  my @months = qw ( jan feb mar apr may jun jul aug sep oct nov dec);
  my %monthNumbers = (
    'jan' => 1,
    'feb' => 2,
    'mar' => 3,
    'apr' => 4,
    'may' => 5,
    'jun' => 6,
    'jul' => 7,
    'aug' => 8,
    'sep' => 9,
    'oct' => 10,
    'nov' => 11,
    'dec' => 12,
  );
  
  my $summary = clean_summary($values->{'Summary'});

  my @words = split(/ +/, lc $summary);
  for (my $i = 0; $i < scalar(@words); $i++) {
    foreach my $clue (@months) {
      if ($words[$i] =~ /^$clue/i) {
        my $month = $monthNumbers{$clue};
        my $day = '';
        if (exists($words[$i + 1]) && $words[$i + 1] =~ /\d+/) {
          $day = $words[$i + 1];
        } else {
          $day = 1;
        }
        return $month . '/' . $day . '/2018';        
      }
    }
  }

  return $startDate;
}

sub get_end_date {
  my $values = shift;

  my $endDate = $values->{'SessionEndDate'};

  return $endDate unless ($endDate eq '12/31/2018');

  # Look for the last "month" word in the summary and use it
  my @months = qw ( jan feb mar apr may jun jul aug sep oct nov dec);
  my %monthNumbers = (
    'jan' => 1,
    'feb' => 2,
    'mar' => 3,
    'apr' => 4,
    'may' => 5,
    'jun' => 6,
    'jul' => 7,
    'aug' => 8,
    'sep' => 9,
    'oct' => 10,
    'nov' => 11,
    'dec' => 12,
  );
  my %lastMonthDays = (
    'jan' => 31,
    'feb' => 28,
    'mar' => 31,
    'apr' => 30,
    'may' => 31,
    'jun' => 30,
    'jul' => 31,
    'aug' => 31,
    'sep' => 30,
    'oct' => 31,
    'nov' => 30,
    'dec' => 31,
  );

  my $summary = clean_summary($values->{'Summary'});

  my @words = split(/ +/, lc $summary);
  for (my $i = $#words; $i >= 0; $i--) {
    foreach my $clue (@months) {
      if ($words[$i] =~ /^$clue/i) {
        my $month = $monthNumbers{$clue};
        my $day = '';
        if (exists($words[$i + 1]) && $words[$i + 1] =~ /\d+/) {
          $day = $words[$i + 1];
        } else {
          $day = $lastMonthDays{$clue};
        }
        return $month . '/' . $day . '/2018';        
      }
    }
  }

  return $endDate;
}

sub clean_summary {
  my $summary = shift;

  $summary =~ s/20(17|18)//g;
  $summary =~ s/6 ?& up//;
  $summary =~ s/[\.\-:,]/ /g;
  $summary =~ s/(\d*)(th|rd|nd|st)/$1/g;
  $summary =~ s/\d*(am|pm)//g;

  return $summary;
}

sub clean_childcare_values {
  my $values = shift;

  $values->{'Source'} = 'childcare';
  
  $values->{'MappedProgramDescription'} = map_childcare_descriptions($values);

  return clean_all_values($values);
}

sub clean_camp_values {
  my $values = shift;

  $values->{'Source'} = 'camp';

  $values->{'ListPrice'} = $values->{'FullMemberPrice'};

  $values->{'ShortPayProcCode'} = 'AR';

  $values->{'Summary'} = $values->{'ClassSummary'};

  my $startDate = ParseDate($values->{'SessionStartDate'} . ' 9:00am');
  $values->{'StartDateTime'} = UnixDate($startDate, '%Y-%m-%d %r');
  $values->{'EndDateTime'} = UnixDate(DateCalc($startDate, '+5 days 8 hours'), '%Y-%m-%d %r');

  $values->{'MappedProgramDescription'} = map_camp_descriptions($values);

  return clean_all_values($values);
}

sub clean_all_values {
  my $values = shift;

  $values->{'BranchCode'} = branch_name_map()->{$values->{'BranchName'}};

  if (!$values->{'BranchCode'}) {
    print "No branch code for $values->{'BranchName'}\n";
    exit;
  }

  $values->{'ShortPayProcCode'} = 'REJECT' unless ($values->{'ShortPayProcCode'});
  $values->{'ClCccFlag'} = 'Y' unless ($values->{'ClCccFlag'});

  $values->{'LastRegistrationDate'} = UnixDate(DateCalc($values->{'EndDateTime'}, 
    '-3 weeks'), '%Y-%m-%d') if ($values->{'EndDateTime'});
  $values->{'LastRefundDate'} = UnixDate(DateCalc($values->{'EndDateTime'}, 
    '+3 months'), '%Y-%m-%d')if ($values->{'EndDateTime'});
  $values->{'ExpirationDate'} = UnixDate(DateCalc($values->{'StartDateTime'}, 
    '+3 weeks'), '%Y-%m-%d') if ($values->{'StartDateTime'});


  foreach my $key (qw( ListPrice NonMemberPrice FullMemberPrice ProgramParticipantPrice)) {
    $values->{$key} =~ s/[\$,]//g;
  }

  $values->{'StartDateTime'} = '' unless ($values->{'StartDateTime'});
  $values->{'EndDateTime'} = '' unless ($values->{'EndDateTime'});
  $values->{'LastRegistrationDate'} = '' unless ($values->{'LastRegistrationDate'});
  $values->{'LastRefundDate'} = '' unless ($values->{'LastRefundDate'});
  $values->{'ExpirationDate'} = '' unless ($values->{'ExpirationDate'});

  $values->{'LimitedSeatsThreshold'} = 10;

  $values->{'MinAgeMonths'} = 0;
  $values->{'MaxAgeMonths'} = 0;

  $values->{'GlAccount'} = format_gl_account($values->{'GlAccount'});
  $values->{'SubsidyGl'} = format_gl_account($values->{'SubsidyGl'});
  ($values->{'PromoDiscountAccount'} = $values->{'GlAccount'}) =~ s/\d{4}$/1396/;
  # Fixed
  $values->{'ArAccount'} = '1-10-10-19-6312';
  $values->{'PplAccount'} = '1-10-10-01-7360';

  # Revenue
  $values->{'WriteOffAccount'} = $values->{'GlAccount'};
  $values->{'CancelAccount'} = $values->{'GlAccount'};
  
  # Promo discount
  $values->{'DiscountAccount'} = $values->{'PromoDiscountAccount'};
  
  # Fixed
  $values->{'DefDiscountAccount'} = '1-10-10-01-7431';

  # Promo discount
  $values->{'AgencyDiscountAccount'} = $values->{'PromoDiscountAccount'};

  # Fixed
  $values->{'AgencyDefDiscountAccount'} = '1-10-10-01-7431';

  # Revenue
  $values->{'RevenueAccount'} = $values->{'GlAccount'};

  # Fixed
  $values->{'DeferredAccount'} = '1-10-10-01-7431';

  if (exists($values->{'Schedule'})) {
    if (grep { $values->{'Schedule'} =~ /$_/i } ('mon-fri', 'mon - fri', 'm-f', 'daily', 'weekly')) {
      $values->{'ClMon'} = 'Y';
      $values->{'ClTue'} = 'Y';
      $values->{'ClWed'} = 'Y';
      $values->{'ClThu'} = 'Y';
      $values->{'ClFri'} = 'Y';
    }

    if ($values->{'Schedule'} =~ /m\/w\/f/i || $values->{'Schedule'} =~ /m-w/i) {
      $values->{'ClMon'} = 'Y';
      $values->{'ClWed'} = 'Y';
      $values->{'ClFri'} = 'Y';
    }

    if ($values->{'Schedule'} =~ /t\/th/i) {
      $values->{'ClTue'} = 'Y';
      $values->{'ClThu'} = 'Y';
    }

    if ($values->{'Schedule'} =~ /m\/w/i) {
      $values->{'ClMon'} = 'Y';
      $values->{'ClWed'} = 'Y';
    }
  }

  foreach my $key (qw(ClMon ClTue ClWed ClThu ClFri ClSat ClSun)) {
    $values->{$key} = 'N' unless ($values->{$key});
  }

  return $values;
}

sub format_gl_account {
  my $account = shift;

  return '' unless ($account && $account =~ /(\d{2})-(\d{2})(\d{2})-(\d{4})/);

  return join('-', '1', $1, $2, $3, $4);
}

sub map_program_descriptions {
  my $values = shift;

  if ($values->{'ProgramType'} eq '2018 Friday Night Out/Plus') {
    my $mappedDescription = 'Friday Night Out';

    $mappedDescription = 'Friday Night Out Plus' if ($values->{'ProgramDescription'} =~ /overnight/i);

    return $mappedDescription;
  }

  if ($values->{'ProgramType'} eq '2018 Adult Sports') {
    my $mappedDescription = $values->{'ProgramDescription'};

    foreach my $clue (qw( Basketball Softball Volleyball )) {
      next unless ($mappedDescription =~ /$clue/i);
      $mappedDescription = 'Adult ' . $clue;
      last;
    }

    return $mappedDescription;
  }

  if (grep { $values->{'ProgramType'} =~ /$_/ } ('2018 Little League', '2018 Summer T-Ball')) {
    return 'Youth Baseball';
  }

  my $mappedDescription = '';

  foreach my $level (qw( Preschool Toddler Youth )) {
    if ($values->{'ProgramDescription'} =~ /$level (Stage )?(\d+)/i) {
      $mappedDescription = "$level Swim Level $2";
      last;
    }
  }

  return $mappedDescription if ($mappedDescription);

  if ($values->{'ProgramDescription'} =~ /Barracuda/i) {
    foreach my $level (qw( Preschool Toddler Youth )) {
      if ($values->{'ItemDescription'} =~ /$level (Stage )?(\d+)/i) {
        $mappedDescription = "$level Swim Level $2";
        last;
      }      
    }
    if ($values->{'ItemDescription'} =~ /School Age (Stage )?(\d+)/i) {
      $mappedDescription = "Youth Swim Level $2";
    }
    
    if ($values->{'ItemDescription'} =~ /Pike (Stage )?I/i) {
      $mappedDescription = "Preschool Swim Level 1";
    }
  }

  return $mappedDescription if ($mappedDescription);

  $mappedDescription = $values->{'ProgramDescription'};

  my %map = (
    '(Adult|Teen) Swim Lesson' => 'Adult/Teen Swim Lesson',
    'Youth.*Soccer' => 'Youth Soccer',
    'Miracle League' => 'Joe Nuxhall Youth Miracle League',
    'Scooter' => 'Scooter',
    '(Preschool|PS) Rollers' => 'Preschool Rollers',
    'Gliders' => 'Youth Gliders',
    'Exhibition Team' => 'Exhibition Team',
    'Step .* Sculpt?' => 'Step and Sculpt',
    'Zumba' => 'Zumba',
    'Rope' => 'Ropes',
    'Total Body Conditioning' => 'Total Body Conditionings',
    'P\.?H\.?I\.?T\.?' => 'PHIT',
    'Healthy Living Program' => 'Healthy Living Program',
    'Studio Cycle' => 'Cycling',
    '2 Hour Cycle' => 'Cycling',
    '(Stage A|Swim Starters A)' => 'Parent/Child Level A',
    '(Stage B|Swim Starters B)' => 'Parent/Child Level B',
    'Lifeguarding' => 'YMCA Lifeguard Training',
    'Lifeguard certification' => 'YMCA Lifeguard Recertification',
    'Babysitting' => 'ASHI Child and Babysitting Safety',
    'ASHI.*CPR' => 'ASHI CPR Basic',
    'YMCA Swim Instructor Training' => 'YMCA Swim Lesson Instructor Training',
    'H2O Boot Camp' => 'Water Boot Camp',
    'Step Water Aerobics' => 'Water Step',
    'Muscle & Joint Class' => 'Muscle and Joint',
    'Tae Kwon Do' => 'Tae Kwon Do',
    'BLS' => 'ASHI Basic Life Support',
    'Super Saturday' => 'Friday Night Out',
    'Thanksgiving' => 'Thanksgiving',
    'Jr.*Leaders.*Club' => 'Jr Leaders Club',
    'H20 Deep' => 'Deep H2O',
    'Cardio Splash' => 'Cardio Splash',
    'Kid ?Fit' => 'Kids in Action/Kid Fit',
    'Lazy' => 'Lazy Man Triathlon',
    'Teen Leaders' => 'Teen Leaders Club',
    'Strength Train Together' => 'Strength Train Together',
    'Defend Together' => 'Defend Together',
    'Pumpkin Dash' => 'Pumpkin Dash',
    'Total Body Conditioning' => 'Total Body Conditioning',
    'Interval Boot' => 'Interval Boot Camp',
    'First Aid.*(Oxygen|02)' => 'ASHI First Aid & Emergency Oxygen',
    'H2O?-Cardio-O' => 'H2O-Cardio-O',
    'Preschool Art camp mini' => 'Pee Wee Mini Arts Camp',
    'Crunch +Time' => 'Crunch Time',
  );
  
  foreach my $clue (keys %map) {
    next unless ($mappedDescription =~ /$clue/i);
    $mappedDescription = $map{$clue};
    last;
  }

  return $mappedDescription;

}

sub map_childcare_descriptions {
  my $values = shift;

  my $mappedDescription = $values->{'ProgramDescription'};

  # Before School Part Time
  # Before School Full Time
  # After School Part Time
  # After School Full Time
  # Before & After School Part Time
  # Before & After School Full Time
  # School's Day Out

  # PT or Ft
  my $schedule = '';
  SCHEDULE: {
    if ($values->{'ProgramDescription'} =~ /(Part Time|PT)/i) {
      $schedule = 'Part Time';
      last SCHEDULE;
    }

    if ($values->{'Schedule'} =~ /(Part Time|PT)/i) {
      $schedule = 'Part Time';
      last SCHEDULE;
    }
  }
  
  my %map = (
    'Day Out' => 'School\'s Day Out',
    'School Out' => 'School\'s Day Out',
  );
  
  foreach my $clue (keys %map) {
    next unless ($mappedDescription =~ /$clue/i);
    $mappedDescription = $map{$clue};
    last;
  }

  return $mappedDescription;

}

sub map_camp_descriptions {
  my $values = shift;

  my $mappedDescription = $values->{'ProgramDescription'};

  my %map = (
    'Adventure Camp' => 'Adventure Camp',
    'Teen Camp' => 'Teen Camp',
    'Preschool' => 'Preschool Fun Camp',
    'Nerf' => 'Nerf Camp',
    'CIT' => 'Counselor In Training',
    'Aquatic' => 'Aquatics Camp',
    'Pee Wee Swim' => 'Pee Wee Aquatic Camp',
    'Water Safety' => 'Water Safety Camp',
    'Flag Football' => 'Flag Football Camp',
    'All Sorts' => 'All Sports Camp',
    'Basketball' => 'Basketball Camp',
    'Lacrosse' => 'Lacrosse Camp',
  );
  
  foreach my $clue (keys %map) {
    next unless ($mappedDescription =~ /$clue/i);
    $mappedDescription = $map{$clue};
    last;
  }

  return $mappedDescription;

}

sub programCodesHeaderMap {
  return {
    "\x{feff}Branch" => 'Branch',
    'Department (2)' => 'DepartmentCode',
    'Dept Name' => 'DepartmentName',
    'SubDepartment (4)' => 'SubDepartmentCode',
    'SubDept Name' => 'SubDepartmentName',
    'Session (2)' => 'Session',
    'Date (6)' => 'Date',
    'Season (3)' => 'Season',
    'Product Class' => 'ProductClass',
    'SubClass' => 'SubClass',
    'Category' => 'Category',
    'SubCategory' => 'SubCategory',
  };
}

sub programsHeaderMap {
  return {
    'CycleName' => 'ProgramType',
    'Description' => 'ProgramDescription',
    'MemberFee' => 'FullMemberPrice',
    'MaximumAge' => 'MaxAgeYears',
    'MaxEnroll' => 'MaxCapacity',
    'MinimumAge' => 'MinAgeYears',
    'NonMemberFee' => 'NonMemberPrice',
    'BasicMemberFee' => 'ProgramParticipantPrice',
  };
}

sub childCareHeaderMap {
  return {
    "\x{feff}Branch" => 'BranchName',
    'Class Max Age' => 'MaxAgeYears',
    'Class Min Age' => 'MinAgeYears',
    'End Date' => 'EndDateTimeString',
    'Full Members' => 'FullMemberPrice',
    'Max Capacity' => 'MaxCapacity',
    'Non-Members' => 'NonMemberPrice',
    'Program Participant' => 'ProgramParticipantPrice',
    'Start Date' => 'SessionStartDate',
    'Status' => 'Active',
    'GL Account' => 'GlAccount',
    'Program Description' => 'ProgramDescription',
    'Program Type' => 'ProgramType',
    'Subsidy GL' => 'SubsidyGl',
  };
}

sub campHeaderMap {
  return {
    "Branch" => 'BranchName',
    'Class Max Age' => 'MaxAgeYears',
    'Class Min Age' => 'MinAgeYears',
    'Full Members' => 'FullMemberPrice',
    'GL Account' => 'GlAccount',
    'Max Capacity' => 'MaxCapacity',
    'Non-Members' => 'NonMemberPrice',
    'Program Description' => 'ProgramDescription',
    'Program Participant' => 'ProgramParticipantPrice',
    'Program Type' => 'ProgramType',
    'Session Start Date' => 'SessionStartDate',
    'Status' => 'Active',
    'Class summary' => 'ClassSummary',
    'Deposit Enable' => 'DepositEnable',
    'Min Deposit' => 'MinDeposit',
    'Subsidy GL' => 'SubsidyGl',
    'Wait List' => 'WaitList',    
  };
}
