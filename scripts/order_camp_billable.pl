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

my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

# Export CL_CUSTOMER Excel file to CSV
#   delete comments row, headings should be row 1
# Save as encoding
# Delete comment row
my $templateName = 'CL_CUSTOMER';

my $columnMap = {
  'CL_CUSTOMER_ID'                          => { 'type' => 'record', 'source' => 'CL_CUSTOMER_ID' },
  'TRX ID'                                  => { 'type' => 'record', 'source' => 'TRX ID' },
  'CL_DRS_CARE_DESC'                        => { 'type' => 'record', 'source' => 'CL_DRS_CARE_DESC' },
  'CL_CONDITIONS_REQUIRED_SPECIAL_CARE'     => { 'type' => 'record', 'source' => 'CL_CONDITIONS_REQUIRED_SPECIAL_CARE' },
  'CL_DESCRIPTION_YES'                      => { 'type' => 'record', 'source' => 'CL_DESCRIPTION_YES' },
  'CL_DIETARY_RESTRICTIONS_DESCRIPTION'     => { 'type' => 'record', 'source' => 'CL_DIETARY_RESTRICTIONS_DESCRIPTION' },
  'CL_MED_DOSAGE_PURPOSE'                   => { 'type' => 'record', 'source' => 'CL_MED_DOSAGE_PURPOSE' },
  'CL_MEDICAL_ALLERGIES_DESCRIPTION'        => { 'type' => 'record', 'source' => 'CL_MEDICAL_ALLERGIES_DESCRIPTION' },
  'CL_SURGERIES_INJURIES_ILLNESS'           => { 'type' => 'record', 'source' => 'CL_SURGERIES_INJURIES_ILLNESS' },
  'CL_TRAVEL_OUTSIDE_12'                    => { 'type' => 'record', 'source' => 'CL_TRAVEL_OUTSIDE_12' },
  'CL_DEFAULT_BILL_TO_FLAG'                 => { 'type' => 'record', 'source' => 'CL_DEFAULT_BILL_TO_FLAG' },
  'CL_DEFAULT_BILL_TO_LABEL_NAME'           => { 'type' => 'record', 'source' => 'CL_DEFAULT_BILL_TO_LABEL_NAME' },
  'CL_DEFAULT_BILL_TO_CUSTOMER_ID'          => { 'type' => 'record', 'source' => 'CL_DEFAULT_BILL_TO_CUSTOMER_ID' },
  'CL_DEFAULT_BILL_TO_CUSTOMER_TRX_ID'      => { 'type' => 'record', 'source' => 'CL_DEFAULT_BILL_TO_CUSTOMER_TRX_ID' },
  'CL_ALERT'                                => { 'type' => 'record', 'source' => 'CL_ALERT' },
  'CL_HEAR_YMCA'                            => { 'type' => 'record', 'source' => 'CL_HEAR_YMCA' },
  'CL_HOUSEHOLD_INCOME'                     => { 'type' => 'record', 'source' => 'CL_HOUSEHOLD_INCOME' },
  'CL_INTERVIEW_DATE'                       => { 'type' => 'record', 'source' => 'CL_INTERVIEW_DATE' },
  'CL_INTERVIEW_GIVEN'                      => { 'type' => 'record', 'source' => 'CL_INTERVIEW_GIVEN' },
  'CL_INTERVIEW_TYPE'                       => { 'type' => 'record', 'source' => 'CL_INTERVIEW_TYPE' },
  'CL_INTERVIEWED_BY'                       => { 'type' => 'record', 'source' => 'CL_INTERVIEWED_BY' },
  'CL_CURRENT_ACTIVITY_LEVEL'               => { 'type' => 'record', 'source' => 'CL_CURRENT_ACTIVITY_LEVEL' },
  'CL_JOIN_DATE'                            => { 'type' => 'record', 'source' => 'CL_JOIN_DATE' },
  'CL_JOIN_REASON'                          => { 'type' => 'record', 'source' => 'CL_JOIN_REASON' },
  'CL_ASSIGNED_TO'                          => { 'type' => 'record', 'source' => 'CL_ASSIGNED_TO' },
  'CL_INDIVIDUAL_OR_FAMILY'                 => { 'type' => 'record', 'source' => 'CL_INDIVIDUAL_OR_FAMILY' },
  'CL_NEW_MEMBER'                           => { 'type' => 'record', 'source' => 'CL_NEW_MEMBER' },
  'CL_ISSUING_BRANCH'                       => { 'type' => 'record', 'source' => 'CL_ISSUING_BRANCH' },
  'CL_PRIMARY_BUSINESS'                     => { 'type' => 'record', 'source' => 'CL_PRIMARY_BUSINESS' },
  'CL_ACTIVITY'                             => { 'type' => 'record', 'source' => 'CL_ACTIVITY' },
  'CL_BECAUSE_HELPING'                      => { 'type' => 'record', 'source' => 'CL_BECAUSE_HELPING' },
  'CL_BECAUSE_MARKETING'                    => { 'type' => 'record', 'source' => 'CL_BECAUSE_MARKETING' },
  'CL_BECAUSE_RESEARCH'                     => { 'type' => 'record', 'source' => 'CL_BECAUSE_RESEARCH' },
  'CL_FACILITY_CARD_NO'                     => { 'type' => 'record', 'source' => 'CL_FACILITY_CARD_NO' },
  'CL_CARD_ISSUEDATE'                       => { 'type' => 'record', 'source' => 'CL_CARD_ISSUEDATE' },
  'CL_PREV_FACILITY_CARDNO'                 => { 'type' => 'record', 'source' => 'CL_PREV_FACILITY_CARDNO' },
  'CL_TOTAL_GUEST_PASS_COUNT'               => { 'type' => 'record', 'source' => 'CL_TOTAL_GUEST_PASS_COUNT' },
  'CL_ACT_MEM_PRODUCT_ID'                   => { 'type' => 'record', 'source' => 'CL_ACT_MEM_PRODUCT_ID' },
  'CL_COLLEGE'                              => { 'type' => 'record', 'source' => 'CL_COLLEGE' },
  'CL_YEAR_IN_SCHOOL'                       => { 'type' => 'record', 'source' => 'CL_YEAR_IN_SCHOOL' },
  'CL_GRADUATION'                           => { 'type' => 'record', 'source' => 'CL_GRADUATION' },
  'CL_ACADEMIC_MAJOR'                       => { 'type' => 'record', 'source' => 'CL_ACADEMIC_MAJOR' },
  'CL_STUDENT_ID'                           => { 'type' => 'record', 'source' => 'CL_STUDENT_ID' },
  'CL_LAST_VISIT'                           => { 'type' => 'record', 'source' => 'CL_LAST_VISIT' },
  'CL_BRANCH'                               => { 'type' => 'record', 'source' => 'CL_BRANCH' },
  'CL_GIVEN_BY'                             => { 'type' => 'record', 'source' => 'CL_GIVEN_BY' },
  'CL_OCINFO_ACKNOWLEDGEMENT'               => { 'type' => 'record', 'source' => 'CL_OCINFO_ACKNOWLEDGEMENT' },
  'CL_DESCRIPTION_ACA'                      => { 'type' => 'record', 'source' => 'CL_DESCRIPTION_ACA' },
  'CL_CLINIC_ADDRESS'                       => { 'type' => 'record', 'source' => 'CL_CLINIC_ADDRESS' },
  'CL_CONDITIONS_SPECIAL_CARE'              => { 'type' => 'record', 'source' => 'CL_CONDITIONS_SPECIAL_CARE' },
  'CL_DENTIST_NAME'                         => { 'type' => 'record', 'source' => 'CL_DENTIST_NAME' },
  'CL_DIETARY_RESTRICTIONS'                 => { 'type' => 'record', 'source' => 'CL_DIETARY_RESTRICTIONS' },
  'CL_DNT_CLINIC_ADDRESS'                   => { 'type' => 'record', 'source' => 'CL_DNT_CLINIC_ADDRESS' },
  'CL_DRS_CARE'                             => { 'type' => 'record', 'source' => 'CL_DRS_CARE' },
  'CL_DTP'                                  => { 'type' => 'record', 'source' => 'CL_DTP' },
  'CL_HEALTH_INSURANCE'                     => { 'type' => 'record', 'source' => 'CL_HEALTH_INSURANCE' },
  'CL_HEPA'                                 => { 'type' => 'record', 'source' => 'CL_HEPA' },
  'CL_HEPB'                                 => { 'type' => 'record', 'source' => 'CL_HEPB' },
  'CL_HIB'                                  => { 'type' => 'record', 'source' => 'CL_HIB' },
  'CL_INS_CARRIER'                          => { 'type' => 'record', 'source' => 'CL_INS_CARRIER' },
  'CL_INS_GRP_NUMBER'                       => { 'type' => 'record', 'source' => 'CL_INS_GRP_NUMBER' },
  'CL_INS_POLICY_ID'                        => { 'type' => 'record', 'source' => 'CL_INS_POLICY_ID' },
  'CL_LAST_PHYSICAL_DATE'                   => { 'type' => 'record', 'source' => 'CL_LAST_PHYSICAL_DATE' },
  'CL_MEDICAL_ALLERGIES'                    => { 'type' => 'record', 'source' => 'CL_MEDICAL_ALLERGIES' },
  'CL_MMR'                                  => { 'type' => 'record', 'source' => 'CL_MMR' },
  'CL_PCV'                                  => { 'type' => 'record', 'source' => 'CL_PCV' },
  'CL_PHYSICIAN_NAME'                       => { 'type' => 'record', 'source' => 'CL_PHYSICIAN_NAME' },
  'CL_POLIO'                                => { 'type' => 'record', 'source' => 'CL_POLIO' },
  'CL_SURGERIES_ILLNESS'                    => { 'type' => 'record', 'source' => 'CL_SURGERIES_ILLNESS' },
  'CL_TAKING_PRESCRIPTIONS'                 => { 'type' => 'record', 'source' => 'CL_TAKING_PRESCRIPTIONS' },
  'CL_TETANUS'                              => { 'type' => 'record', 'source' => 'CL_TETANUS' },
  'CL_TRAVEL_12'                            => { 'type' => 'record', 'source' => 'CL_TRAVEL_12' },
  'CL_VARICELLA'                            => { 'type' => 'record', 'source' => 'CL_VARICELLA' },
  'CL_CLINIC_ADDRESS1'                      => { 'type' => 'record', 'source' => 'CL_CLINIC_ADDRESS1' },
  'CL_DENTIST_CITY'                         => { 'type' => 'record', 'source' => 'CL_DENTIST_CITY' },
  'CL_DENTIST_STATE'                        => { 'type' => 'record', 'source' => 'CL_DENTIST_STATE' },
  'CL_DENTIST_ZIP'                          => { 'type' => 'record', 'source' => 'CL_DENTIST_ZIP' },
  'CL_PHYSICIAN_CITY'                       => { 'type' => 'record', 'source' => 'CL_PHYSICIAN_CITY' },
  'CL_PHYSICIAN_STATE'                      => { 'type' => 'record', 'source' => 'CL_PHYSICIAN_STATE' },
  'CL_PHYSICIAN_ZIP'                        => { 'type' => 'record', 'source' => 'CL_PHYSICIAN_ZIP' },
  'CL_DNT_CLINIC_ADDRESS1'                  => { 'type' => 'record', 'source' => 'CL_DNT_CLINIC_ADDRESS1' },
  'CL_INS_CARRIER_PHONE'                    => { 'type' => 'record', 'source' => 'CL_INS_CARRIER_PHONE' },
  'CL_DNT_CLINIC_PHONE'                     => { 'type' => 'record', 'source' => 'CL_DNT_CLINIC_PHONE' },
  'CL_PHYSICIAN_PHONE'                      => { 'type' => 'record', 'source' => 'CL_PHYSICIAN_PHONE' },
  'CL_REL_PARTICIANT'                       => { 'type' => 'record', 'source' => 'CL_REL_PARTICIANT' },
  'CL_INSURED_NAME'                         => { 'type' => 'record', 'source' => 'CL_INSURED_NAME' },
  'CL_WAIVER_DATE'                          => { 'type' => 'record', 'source' => 'CL_WAIVER_DATE' },
  'CL_HEALTH_HISTORY_DATE'                  => { 'type' => 'record', 'source' => 'CL_HEALTH_HISTORY_DATE' },
  'CL_AUTH_DATE'                            => { 'type' => 'record', 'source' => 'CL_AUTH_DATE' },
  'CL_AUTH_ACTIVITY_DATE'                   => { 'type' => 'record', 'source' => 'CL_AUTH_ACTIVITY_DATE' },
  'CL_WAIVER_BRANCH'                        => { 'type' => 'record', 'source' => 'CL_WAIVER_BRANCH' },
  'CL_AUTH_ACTIVITY_BRANCH'                 => { 'type' => 'record', 'source' => 'CL_AUTH_ACTIVITY_BRANCH' },
  'CL_AUTH_BRANCH'                          => { 'type' => 'record', 'source' => 'CL_AUTH_BRANCH' },
  'CL_HEALTH_HISTORY_BRANCH'                => { 'type' => 'record', 'source' => 'CL_HEALTH_HISTORY_BRANCH' },
  'CL_MAIDENNAME'                           => { 'type' => 'record', 'source' => 'CL_MAIDENNAME' },
  'CL_TOURED_BY'                            => { 'type' => 'record', 'source' => 'CL_TOURED_BY' },
  'CL_TOURED_DATE'                          => { 'type' => 'record', 'source' => 'CL_TOURED_DATE' },
  'CL_EMPLOYEE_NO'                          => { 'type' => 'record', 'source' => 'CL_EMPLOYEE_NO' },
  'CL_EMPLOYEE_TERMDATE'                    => { 'type' => 'record', 'source' => 'CL_EMPLOYEE_TERMDATE' },
  'CL_SALE_CAMPHIST_TRIP_TYPE'              => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_TRIP_TYPE' },
  'CL_SALE_CAMPHIST_TRIP_TYPE_PARTNER'      => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_TRIP_TYPE_PARTNER' },
  'CL_SALE_CAMPHIST_VOL_YEAR'               => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_VOL_YEAR' },
  'CL_SALE_CAMPHIST_VOLUNTEER'              => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_VOLUNTEER' },
  'CL_SALE_CAMPHIST_YEAR'                   => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_YEAR' },
  'CL_SALE_CAMPHIST_YPHISTORY'              => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_YPHISTORY' },
  'CL_SALE_CAMPHIST_ALUMINI'                => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_ALUMINI' },
  'CL_SALE_CAMPHIST_BOARD'                  => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_BOARD' },
  'CL_SALE_CAMPHIST_BOARD_YEAR'             => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_BOARD_YEAR' },
  'CL_SALE_CAMPHIST_CALLVOL'                => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_CALLVOL' },
  'CL_SALE_CAMPHIST_CAMPER_YR'              => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_CAMPER_YR' },
  'CL_SALE_CAMPHIST_CAMPER_YR_PARTNER'      => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_CAMPER_YR_PARTNER' },
  'CL_SALE_CAMPHIST_CAPHISTORY'             => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_CAPHISTORY' },
  'CL_SALE_CAMPHIST_DIVISION'               => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_DIVISION' },
  'CL_SALE_CAMPHIST_FWS_STAFF'              => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_FWS_STAFF' },
  'CL_SALE_CAMPHIST_FWS_STAFF_YEAR'         => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_FWS_STAFF_YEAR' },
  'CL_SALE_CAMPHIST_FWS_STAFF_YEAR_PARTNER' => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_FWS_STAFF_YEAR_PARTNER' },
  'CL_SALE_CAMPHIST_NAME'                   => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_NAME' },
  'CL_SALE_CAMPHIST_PARENT'                 => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_PARENT' },
  'CL_SALE_CAMPHIST_PARENT_YEAR'            => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_PARENT_YEAR' },
  'CL_SALE_CAMPHIST_PARENT_YEAR_PARTNER'    => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_PARENT_YEAR_PARTNER' },
  'CL_SALE_CAMPHIST_STAFF_YEAR'             => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_STAFF_YEAR' },
  'CL_SALE_CAMPHIST_STAFF_YEAR_PARTNER'     => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_STAFF_YEAR_PARTNER' },
  'CL_SALE_CAMPHIST_SUMMER_STAFF_ALUMINI'   => { 'type' => 'record', 'source' => 'CL_SALE_CAMPHIST_SUMMER_STAFF_ALUMINI' },
  'CL_FITNESS_EVAL'                         => { 'type' => 'record', 'source' => 'CL_FITNESS_EVAL' },
  'CL_OVERRIDE_AGEING_OUT_FLAG'             => { 'type' => 'record', 'source' => 'CL_OVERRIDE_AGEING_OUT_FLAG' },
  'CL_QUAL_IMMUNIZATION_WAIVER'             => { 'type' => 'record', 'source' => 'CL_QUAL_IMMUNIZATION_WAIVER' },
  'CL_CONNECTOR'                            => { 'type' => 'record', 'source' => 'CL_CONNECTOR' },
  'CL_JOIN_CLASS_MONTH'                     => { 'type' => 'record', 'source' => 'CL_JOIN_CLASS_MONTH' },
  'CL_JOIN_CLASS_YEAR'                      => { 'type' => 'record', 'source' => 'CL_JOIN_CLASS_YEAR' },
  'CL_CONNECTOR2'                           => { 'type' => 'record', 'source' => 'CL_CONNECTOR2' },
  'CL_HLTHINCENT_INS_CO'                    => { 'type' => 'record', 'source' => 'CL_HLTHINCENT_INS_CO' },
  'CL_HLTHINCENT_INS_ID'                    => { 'type' => 'record', 'source' => 'CL_HLTHINCENT_INS_ID' },
  'CL_HLTHINCENT_SIGNUP_DATE'               => { 'type' => 'record', 'source' => 'CL_HLTHINCENT_SIGNUP_DATE' },
  'CL_HLTHINCENT_SIGNUP_BRANCH'             => { 'type' => 'record', 'source' => 'CL_HLTHINCENT_SIGNUP_BRANCH' },
  'CL_HLTHINCENT_GROUP'                     => { 'type' => 'record', 'source' => 'CL_HLTHINCENT_GROUP' },
  'CL_HLTHINCENT_MBR_ID'                    => { 'type' => 'record', 'source' => 'CL_HLTHINCENT_MBR_ID' },
  'CL_CLINIC_PHONE'                         => { 'type' => 'record', 'source' => 'CL_CLINIC_PHONE' },
  'CL_DRLICENSE_NUMBER'                     => { 'type' => 'record', 'source' => 'CL_DRLICENSE_NUMBER' },
  'CL_DRLICENSE_STATE'                      => { 'type' => 'record', 'source' => 'CL_DRLICENSE_STATE' },
  'CL_DRLICENSE_EXP'                        => { 'type' => 'record', 'source' => 'CL_DRLICENSE_EXP' },
  'CL_CONNECT_BRANCH'                       => { 'type' => 'record', 'source' => 'CL_CONNECT_BRANCH' },
  'CL_CHILD_PICKUP_CODE_WORD'               => { 'type' => 'record', 'source' => 'CL_CHILD_PICKUP_CODE_WORD' },
  'CL_CHILD_PICKUP_CODE_LAST_UPDATED'       => { 'type' => 'record', 'source' => 'CL_CHILD_PICKUP_CODE_LAST_UPDATED' },
  'CL_CHILD_PICKUP_CODE_LAST_UPDATED_BY'    => { 'type' => 'record', 'source' => 'CL_CHILD_PICKUP_CODE_LAST_UPDATED_BY' },
  'CL_MBR_CANCEL_REASON'                    => { 'type' => 'record', 'source' => 'CL_MBR_CANCEL_REASON' },
  'CL_MBR_CANCEL_DATE'                      => { 'type' => 'record', 'source' => 'CL_MBR_CANCEL_DATE' },
  'CL_EXIT_INTERVIEW_BY'                    => { 'type' => 'record', 'source' => 'CL_EXIT_INTERVIEW_BY' },
  'CL_EXIT_INTERVIEW_DATE'                  => { 'type' => 'record', 'source' => 'CL_EXIT_INTERVIEW_DATE' },
  'CL_LAST_DRAFT_DATE'                      => { 'type' => 'record', 'source' => 'CL_LAST_DRAFT_DATE' },
  'CL_MBR_CANCEL_COMMENTS'                  => { 'type' => 'record', 'source' => 'CL_MBR_CANCEL_COMMENTS' },
  'CL_JOIN_YEAR'                            => { 'type' => 'record', 'source' => 'CL_JOIN_YEAR' },
  'CL_JOIN_MONTH'                           => { 'type' => 'record', 'source' => 'CL_JOIN_MONTH' },
  'CL_INTERVIEW_COMMENTS'                   => { 'type' => 'record', 'source' => 'CL_INTERVIEW_COMMENTS' },
  'CL_HEALTH_UD1'                           => { 'type' => 'record', 'source' => 'CL_HEALTH_UD1' },
  'CL_HEALTH_UD1_DESC'                      => { 'type' => 'record', 'source' => 'CL_HEALTH_UD1_DESC' },
};

my $csv = Text::CSV_XS->new ({ auto_diag => 1, eol => $/ });

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $orderBillables = {};
my($rateFile, $headers, $totalRows) = open_data_file('data/camp_orders_billing.csv');
while(my $rowIn = $csv->getline($rateFile)) {
  my $values = map_values($headers, $rowIn);
  # use the first one
  next if (exists($orderBillables->{$values->{'personify_id'}}));
  next if ($values->{'billable_per_id'} eq $values->{'personify_id'});
  $orderBillables->{$values->{'personify_id'}} = $values;
}
close($rateFile);

my $row = 1;
process_data_file(
  'data/CL_CUSTOMER.csv',
  sub {
    my $values = shift;
    # dd($values);

    $values->{'CL_CUSTOMER_ID'} = lookup_id($values->{'TRX ID'});

    if ($values->{'CL_DEFAULT_BILL_TO_CUSTOMER_TRX_ID'}) {
      $values->{'CL_DEFAULT_BILL_TO_CUSTOMER_ID'} = lookup_id($values->{'CL_DEFAULT_BILL_TO_CUSTOMER_TRX_ID'});
    }

    # skip if default billable is set
    if (!$values->{'CL_DEFAULT_BILL_TO_CUSTOMER_ID'} && exists($orderBillables->{$values->{'CL_CUSTOMER_ID'}})) {
      my $orderBilling = $orderBillables->{$values->{'CL_CUSTOMER_ID'}};

      $values->{'CL_DEFAULT_BILL_TO_FLAG'} = 'Y';
      $values->{'CL_DEFAULT_BILL_TO_CUSTOMER_ID'} = $orderBilling->{'billable_per_id'};
      $values->{'CL_DEFAULT_BILL_TO_CUSTOMER_TRX_ID'} = $orderBilling->{'billable_trx_id'};
      ($values->{'CL_DEFAULT_BILL_TO_LABEL_NAME'}) = $dbh->selectrow_array(q{
        select c_name
          from name_map
          where p_id = ?
        }, undef, $orderBilling->{'billable_per_id'});

    }

    write_record(
      $worksheet,
      $row++,
      make_record($values, \@allColumns, $columnMap)
    );
  }
);

