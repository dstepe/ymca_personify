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
  'PRICE'                                    => { 'type' => 'record', 'source' => 'NonMemberPrice' },
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

my($dataFile, $headers, $totalRows) = open_data_file('data/ProgramCodes.csv', programCodesHeaderMap());

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
  
  next unless ($values->{'Active'} eq 'YES');
  next if ($values->{'ProgramType'} eq '2018 Spring');

  push(@{$products}, $values);
}

close($dataFile);

# ($dataFile, $headers, $totalRows) = open_data_file('data/Childcare.csv', childCareHeaderMap());

# print "Processing child care\n";
# $progress = Term::ProgressBar->new({ 'count' => $totalRows });

# $count = 1;
# while(my $rowIn = $csv->getline($dataFile)) {

#   $progress->update($count++);

#   my $values = clean_childcare_values(map_values($headers, $rowIn));
#   # dump($values); exit;
#   push(@{$products}, $values);
# }

# close($dataFile);

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
  ['Source', 'Type', 'Description', 'Session Start Date', 'Schedule']);

my $collector = {};
print "Generating program files\n";
our $partTracker = {};
my $availableDate = UnixDate(ParseDate('1/1/2000'), '%Y-%m-%d');

$progress = Term::ProgressBar->new({ 'count' => scalar(@{$products}) });
$count = 1;
my $row = 1;
my $programTypeRow = 1;
foreach my $program (@{$products}) {
  $progress->update($count++);

  $program->{'AvailableDate'} = $availableDate;
  $program->{'ExpirationDate'} = $program->{'SessionEndDate'} || '';
  
  my $productDetails = get_product_details($program);

  $program->{'ProductCode'} = $productDetails->{'ProductCode'};
  $program->{'Department'} = $productDetails->{'Department'};
  $program->{'DepartmentSubClass'} = $productDetails->{'DepartmentSubClass'};

  write_record($programTypeWorksheet, $programTypeRow++, [
    $program->{'Source'},
    $program->{'ProgramType'} || '',
    $program->{'ProgramDescription'} || '',
    $program->{'SessionStartDate'} || '',
    $program->{'Schedule'} || '',
  ]) unless ($program->{'ProductCode'});

  # $collector->{$program->{'ProgramType'} . ' / ' . $program->{'ProgramDescription'}}++;

  # unless (
  #     $program->{'ClMon'} eq 'Y' ||
  #     $program->{'ClTue'} eq 'Y' ||
  #     $program->{'ClWed'} eq 'Y' ||
  #     $program->{'ClThu'} eq 'Y' ||
  #     $program->{'ClFri'} eq 'Y' ||
  #     $program->{'ClSat'} eq 'Y' ||
  #     $program->{'ClSun'} eq 'Y'
  #   ) {
  #   $collector->{$program->{'Schedule'}}++ if (exists($program->{'Schedule'}));
  # }

  my $productRecord = make_record($program, \@productAllColumns, $productColumnMap);
  write_record($productWorksheet, $row, $productRecord);

  my $addRateCodeRecord = make_record($program, \@addRateCodeAllColumns, $addRateCodeColumnMap);
  write_record($addRateCodeWorksheet, $row, $addRateCodeRecord);

  my $addPriceRecord = make_record($program, \@addPriceAllColumns, $addPriceColumnMap);
  write_record($addPriceWorksheet, $row, $addPriceRecord);

  $row++;
}

dump($collector) if (keys %{$collector});

sub get_product_details {
  my $program = shift;
  
  our $programCodes;

  my $productDetails = {
    'ProductCode' => '',
    'Department' => '',
    'DepartmentSubClass' => '',
  };

  return $productDetails unless (
      exists($programCodes->{lc $program->{'MappedProgramDescription'}}) &&
      $program->{'SessionStartDate'}
    );

  my $codeInfo = $programCodes->{lc $program->{'MappedProgramDescription'}};
  
  $productDetails->{'Department'} = $codeInfo->{'ProductClass'};
  $productDetails->{'DepartmentSubClass'} = $codeInfo->{'SubClass'};

  my @codeParts;
  push(@codeParts, $program->{'BranchCode'});
  push(@codeParts, $productDetails->{'Department'});
  push(@codeParts, $productDetails->{'DepartmentSubClass'});
  push(@codeParts, sprintf('%02s', get_program_increment($program->{'BranchCode'}, $codeInfo->{'ProductClass'})));
  push(@codeParts, UnixDate($program->{'SessionStartDate'}, '%m%d%y'));
  push(@codeParts, get_program_season($program->{'ProgramType'}));
  
  $productDetails->{'ProductCode'} = join('_', @codeParts);

  return $productDetails;
}

sub get_program_season {
  my $programType = shift;

  my $season = 'AYR';

  $programType =~ s/^2018 //;

  $season = 'SPR' if ($programType =~ /^Spring/);
  $season = 'SM1' if ($programType =~ /^Summer 1/);

  return $season;
}

sub get_program_increment {
  my $branchCode = shift;
  my $productClass = shift || 'un';

  our $partTracker;

  return ++$partTracker->{$branchCode}{$productClass};
}

sub clean_program_values {
  my $values = shift;

  $values->{'Source'} = 'program';

  $values->{'ClCccFlag'} = 'N';

  $values->{'SessionStartDate'} =~ s/ .*$//;
  $values->{'SessionEndDate'} =~ s/ .*$//;
  my $dow = substr(lc $values->{'WeekDaysString'}, 0, 3);
  if (
      $values->{'SessionStartDate'} && 
      $values->{'SessionStartDate'} !~ /^1\/1\/2018/ &&
      grep { $dow eq $_ } qw( mon tue wed thu fri sat sun )
    ) {

    # Date::Manip doesn't seem to like high week durations.
    $values->{'ClassDuration'} = '140 days' if ($values->{'ClassDuration'} eq '20 weeks');

    $values->{'StartDateTime'} = UnixDate(
      Date_GetNext(
        $values->{'SessionStartDate'} . ' ' . $values->{'ClassStartTime'}, 
        $dow, 
        1
      ), '%Y-%m-%d %r');
    
    $values->{'EndDateTime'} = UnixDate(DateCalc($values->{'StartDateTime'}, 
      '+' . $values->{'ClassDuration'}), '%Y-%m-%d %r');

  }

  $values->{'ClMon'} = 'Y' if ($values->{'WeekDaysString'} =~ /Mon/i);
  $values->{'ClTue'} = 'Y' if ($values->{'WeekDaysString'} =~ /Tue/i);
  $values->{'ClWed'} = 'Y' if ($values->{'WeekDaysString'} =~ /Wed/i);
  $values->{'ClThu'} = 'Y' if ($values->{'WeekDaysString'} =~ /Thu/i);
  $values->{'ClFri'} = 'Y' if ($values->{'WeekDaysString'} =~ /Fri/i);
  $values->{'ClSat'} = 'Y' if ($values->{'WeekDaysString'} =~ /Sat/i);
  $values->{'ClSun'} = 'Y' if ($values->{'WeekDaysString'} =~ /Sun/i);

  $values->{'MappedProgramDescription'} = map_program_descriptions($values);

  return clean_all_values($values);
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
  
  $values->{'ShortPayProcCode'} = 'AR';

  my $startDate = ParseDate($values->{'SessionStartDate'});
  $values->{'StartDateTime'} = UnixDate($startDate, '%Y-%m-%d %r');
  $values->{'EndDateTime'} = UnixDate(DateCalc($startDate, '+5 days'), '%Y-%m-%d');

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

  if ($values->{'StartDateTime'}) {
    $values->{'LastRegistrationDate'} = UnixDate(DateCalc($values->{'SessionStartDate'}, 
      '-3 days'), '%Y-%m-%d');
    $values->{'LastRefundDate'} = UnixDate(DateCalc($values->{'SessionStartDate'}, 
      '+7 days'), '%Y-%m-%d');
  }

  foreach my $key (qw( NonMemberPrice FullMemberPrice ProgramParticipantPrice)) {
    $values->{$key} =~ s/\$//;
  }

  $values->{'StartDateTime'} = '' unless ($values->{'StartDateTime'});
  $values->{'EndDateTime'} = '' unless ($values->{'EndDateTime'});
  $values->{'LastRegistrationDate'} = '' unless ($values->{'LastRegistrationDate'});
  $values->{'LastRefundDate'} = '' unless ($values->{'LastRefundDate'});

  $values->{'LimitedSeatsThreshold'} = 10;

  $values->{'MinAgeMonths'} = 0;
  $values->{'MaxAgeMonths'} = 0;

  $values->{'ArAccount'} = '1-10-10-19-6312';
  $values->{'PplAccount'} = '1-10-10-01-7360';
  $values->{'WriteOffAccount'} = '1-10-60-05-1307';
  $values->{'CancelAccount'} = '1-10-60-05-1307';
  $values->{'DiscountAccount'} = '1-10-60-05-1396';
  $values->{'DefDiscountAccount'} = '1-10-10-01-7431';
  $values->{'AgencyDiscountAccount'} = '1-10-60-05-1396';
  $values->{'AgencyDefDiscountAccount'} = '1-10-10-01-7431';
  $values->{'RevenueAccount'} = '1-10-60-05-1307';
  $values->{'DeferredAccount'} = '1-10-10-19-7430';

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

  if (grep { $_ eq $values->{'ProgramType'} } ('2018 Little League', '2018 Summer T-Ball')) {
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
    "\x{feff}Program No" => 'ProgramNumber',
    'Active?' => 'Active',
    'Branch' => 'BranchName',
    'Class Close Status' => 'CloseStatus',
    'Class Duration' => 'ClassDuration',
    'Class Time' => 'ClassStartTime',
    'Cycle Name' => 'ProgramType',
    'Description' => 'ProgramDescription',
    'Full Members' => 'FullMemberPrice',
    'GL Account' => 'GlAccount',
    'Item Description' => 'ItemDescription',
    'Item End Date' => 'SessionEndDate',
    'Item Start Date' => 'SessionStartDate',
    'Max Age' => 'MaxAgeYears',
    'MAX Enroll' => 'MaxCapacity',
    'Min Age' => 'MinAgeYears',
    'Non-Members' => 'NonMemberPrice',
    'NonMember Enrollment?' => 'NonMemberEnrollment',
    'Price' => 'ListPrice',
    'Program Participant' => 'ProgramParticipantPrice',
    'Scholarship GL Account' => 'ScholarshipGlAccount',
    'Tax Rate' => 'TaxRate',
    'Week Days' => 'WeekDaysString',
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
    "\x{feff}Branch" => 'BranchName',
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
