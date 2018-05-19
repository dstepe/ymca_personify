#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::CSV;
use Date::Manip;
use Text::Table;

my $templateName = 'DCT_CUS_INDIVIDUAL';

my $columnMap = {
  'ORG_ID'                               => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                          => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'TRX_ID'                               => { 'type' => 'record', 'source' => 'MemberId' },
  'NAME_PREFIX'                          => { 'type' => 'record', 'source' => 'Prefix' },
  'FIRST_NAME'                           => { 'type' => 'record', 'source' => 'FirstName' },
  'LAST_NAME'                            => { 'type' => 'record', 'source' => 'LastName' },
  'NAME_SUFFIX'                          => { 'type' => 'record', 'source' => 'Suffix' },
  'NAME_CREDENTIALS'                     => { 'type' => 'record', 'source' => 'GMVYMCA' },
  'NICKNAME'                             => { 'type' => 'record', 'source' => 'CasualName' },
  'CUSTOMER_CLASS_CODE'                  => { 'type' => 'static', 'source' => 'INDIV' },
  'CUSTOMER_STATUS_CODE'                 => { 'type' => 'static', 'source' => 'ACTIVE' },
  'CUSTOMER_STATUS_DATE'                 => { 'type' => 'record', 'source' => 'CreatedDate' },
  'GENDER_CODE'                          => { 'type' => 'record', 'source' => 'Gender' },
  'BIRTH_DATE'                           => { 'type' => 'record', 'source' => 'DateOfBirth' },
  'ADDRESS_1'                            => { 'type' => 'record', 'source' => 'PrimaryAddress1' },
  'ADDRESS_2'                            => { 'type' => 'record', 'source' => 'PrimaryAddress2' },
  'CITY'                                 => { 'type' => 'record', 'source' => 'PrimaryCity' },
  'STATE'                                => { 'type' => 'record', 'source' => 'PrimaryState' },
  'POSTAL_CODE'                          => { 'type' => 'record', 'source' => 'PrimaryZip' },
  'COUNTRY_CODE'                         => { 'type' => 'record', 'source' => 'PrimaryCountry' },
  'ADDRESS_TYPE_CODE'                    => { 'type' => 'static', 'source' => 'HOME' },
  'ADDRESS_STATUS_CODE'                  => { 'type' => 'static', 'source' => 'GOOD' },
  'PHONE_AREA_CODE'                      => { 'type' => 'record', 'source' => 'HomePhone' },
  'PRIMARY_PHONE'                        => { 'type' => 'record', 'source' => 'HomePhone' },
  'PRIMARY_PHONE_LOCATION_CODE'          => { 'type' => 'static', 'source' => 'HOME' },
  'PRIMARY_EMAIL_ADDRESS'                => { 'type' => 'record', 'source' => 'Email' },
  'PRIMARY_EMAIL_LOCATION_CODE'          => { 'type' => 'static', 'source' => 'HOME' },
  'ALLOW_PHONE_FLAG'                     => { 'type' => 'static', 'source' => 'Y' },
  'ALLOW_FAX_FLAG'                       => { 'type' => 'static', 'source' => 'Y' },
  'ALLOW_EMAIL_FLAG'                     => { 'type' => 'static', 'source' => 'Y' },
  'ALLOW_SOLICITATION_FLAG'              => { 'type' => 'static', 'source' => 'Y' },
  'PUBLISH_PRIMARY_PHONE_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'PUBLISH_PRIMARY_FAX_FLAG'             => { 'type' => 'static', 'source' => 'N' },
  'PUBLISH_PRIMARY_EMAIL_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'PUBLISH_URL_FLAG'                     => { 'type' => 'static', 'source' => 'N' },
  'CAN_PLACE_ORDER_FLAG'                 => { 'type' => 'static', 'source' => 'Y' },
  'DO_NOT_CALL_FLAG'                     => { 'type' => 'static', 'source' => 'N' },
  'CURRENCY_CODE'                        => { 'type' => 'static', 'source' => 'USD' },
  'WEB_MOBILE_DIRECTORY_FLAG'            => { 'type' => 'static', 'source' => 'N' },
  'COMM_WEB_MOBILE_DIRECTORY_FLAG'       => { 'type' => 'static', 'source' => 'N' },
  'INCLUDE_IN_WEB_MOBILE_DIRECTORY_FLAG' => { 'type' => 'static', 'source' => 'N' },
  'FORMAL_SALUTATION'                    => { 'type' => 'record', 'source' => 'FormalName' },
  'INFORMAL_SALUTATION'                  => { 'type' => 'record', 'source' => 'FirstName' },
  'ONE_TIME_USE_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'CONFIDENTIAL_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'DIRECTORY_FLAG'                       => { 'type' => 'static', 'source' => 'Y' },
  'DIRECTORY_PRIORITY'                   => { 'type' => 'static', 'source' => '0' },
  'ALLOW_LABEL_SALES_FLAG'               => { 'type' => 'static', 'source' => 'Y' },
  'ALLOW_INTERNAL_MAIL_FLAG'             => { 'type' => 'static', 'source' => 'Y' },
  'BILL_PRIMARY_EMPLOYER_FLAG'           => { 'type' => 'static', 'source' => 'N' },
  'TAXABLE_FLAG'                         => { 'type' => 'static', 'source' => 'Y' },
  'EXHIBITOR_FLAG'                       => { 'type' => 'static', 'source' => 'N' },
  'SPEAKER_FLAG'                         => { 'type' => 'static', 'source' => 'N' },
  'SPK_ALLOW_PUBLISH_FLAG'               => { 'type' => 'static', 'source' => 'N' },
  'SPK_ALLOW_RECORD_FLAG'                => { 'type' => 'static', 'source' => 'N' },
  'SPK_ALLOW_PHOTOGRAPH_FLAG'            => { 'type' => 'static', 'source' => 'N' },
  'SPK_ALLOW_INTERVIEW_FLAG'             => { 'type' => 'static', 'source' => 'N' },
  'DONOR_FLAG'                           => { 'type' => 'static', 'source' => 'N' },
  'FOUNDATION_FLAG'                      => { 'type' => 'static', 'source' => 'N' },
  'FND_MATCHING_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'ALLOW_ADVOCACY_FLAG'                  => { 'type' => 'static', 'source' => 'N' },
  'ALLOW_SYSTEM_NOTIFICATION_FLAG'       => { 'type' => 'static', 'source' => 'Y' },
  'ANONYMOUS_FLAG'                       => { 'type' => 'static', 'source' => 'N' },
  'FAMILY_FLAG'                          => { 'type' => 'static', 'source' => 'N' },
  'SOLICITOR_FLAG'                       => { 'type' => 'static', 'source' => 'N' },
  'SOLICITOR_ACTIVE_FLAG'                => { 'type' => 'static', 'source' => 'N' },
  'PRIMARY_SEARCH_GROUP_OVERRIDE_FLAG'   => { 'type' => 'static', 'source' => 'N' },
  'GUEST_CHECKOUT_FLAG'                  => { 'type' => 'static', 'source' => 'N' },
  'PRIMARY_MOBILE_PHONE'                 => { 'type' => 'static', 'source' => 'CellPhone' },
  'PRIMARY_MOBILE_PHONE_LOCATION_CODE'   => { 'type' => 'static', 'source' => 'HOME' },
  'PUBLISH_PRIMARY_MOBILE_PHONE_FLAG'    => { 'type' => 'static', 'source' => 'N' },
};

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $csv = Text::CSV->new();

$/ = "\r\n";

my($membersFile, $headers) = openMembersFile();

my $members = {};
my $families = {};
my $familyOldestMember = {};
my $conflicts = [];
my $noFamily = [];
my $count = 0;
while(my $line = <$membersFile>) {
  chomp $line;

  $count++;
  print "Count $count\n" if ($count % 1000 == 0);
  next unless ($line =~ /F193819563/);

  $csv->parse($line) || die "Line could not be parsed: $line";

  my $values = map_values($headers, [$csv->fields()]);
  next unless ($values->{'FamilyId'} eq 'F193819563');
  # print Dumper($values); exit;

  $members->{$values->{'MemberId'}} = $values;

  unless ($values->{'FamilyId'}) {
    push(@{$noFamily}, {
      'MemberId' => $values->{'MemberId'},
    });
  }
  my $familyId = $values->{'FamilyId'};

  unless (exists($families->{$familyId})) {
    $families->{$familyId} = {
      'primaryId' => '',
      'members' => [],
      'address' => {
        'address1' => '',
        'address2' => '',
        'city' => '',
        'state' => '',
        'zip' => ''
      },
      'email' => '',
      'nextBillDate' => '',
    };
  }

  if ($values->{'DateOfBirth'}) {
    my $memberBirthDate = ParseDate($values->{'DateOfBirth'});

    if (!exists($familyOldestMember->{$familyId})) {
      $familyOldestMember->{$familyId} = {
        'birthDate' => $memberBirthDate,
        'memberId' => $values->{'MemberId'},
      };
    } elsif (Date_Cmp($memberBirthDate, $familyOldestMember->{$familyId}{'birthDate'}) == -1) {
      $familyOldestMember->{$familyId}{'birthDate'} = $memberBirthDate;
      $familyOldestMember->{$familyId}{'memberId'} = $values->{'MemberId'};
    }
  }

  push(@{$families->{$familyId}{'members'}}, $values->{'MemberId'});
  
  if (billableMember($values)) {
    if ($families->{$familyId}{'primaryId'} && 
        $families->{$familyId}{'primaryId'} ne $values->{'MemberId'}) {
      #compare($values, $members->{$families->{$familyId}{'primaryId'}});
      my $conflictedMember = $members->{$families->{$familyId}{'primaryId'}};
      push(@{$conflicts}, {
        'familyId' => $familyId,
        'a-memberId' => $values->{'MemberId'},
        'a-billableId' => $values->{'BillableMemberId'},
        'a-membershipType' => $values->{'MembershipType'},
        'b-memberId' => $conflictedMember->{'MemberId'},
        'b-billableId' => $conflictedMember->{'BillableMemberId'},
        'b-membershipType' => $conflictedMember->{'MembershipType'},
      });
      next;
    } 

    $families->{$familyId}{'primaryId'} = $values->{'MemberId'};
    $families->{$familyId}{'nextBillDate'} = $values->{'NextBillDate'};
    $families->{$familyId}{'email'} = $values->{'Email'};
    $families->{$familyId}{'address'}{'address1'} = $values->{'Address1'};
    $families->{$familyId}{'address'}{'address2'} = $values->{'Address2'};
    $families->{$familyId}{'address'}{'city'} = $values->{'City'};
    $families->{$familyId}{'address'}{'state'} = $values->{'State'};
    $families->{$familyId}{'address'}{'zip'} = $values->{'Zip'};
  }
}

close($membersFile);

# Find families with no primary and use oldest member
my $primaryByBillable = 0;
my $primaryByBirth = 0;
my $missingPrimary = 0;
my $familyCount = 0;
foreach my $familyId (keys %{$families}) {
  $familyCount++;

  if ($families->{$familyId}{'primaryId'}) {
    $primaryByBillable++;
  } else {
    if (exists($familyOldestMember->{$familyId})) {
      $families->{$familyId}{'primaryId'}
        = $familyOldestMember->{$familyId}{'memberId'};
      $primaryByBirth++;
    }
  }

  $missingPrimary++ unless ($families->{$familyId}{'primaryId'});
}

print "families: $familyCount\n";
print "primary by billable: $primaryByBillable\n";
print "primary by birth: $primaryByBirth\n";
print "no primary: $missingPrimary\n";

CONFLICTS: {
  my $conflictWorkbook = make_workbook('conflicted_primary');
  my $conflictWorksheet = make_worksheet($conflictWorkbook, 
    ['FamilyId', 'A MemberId', 'A BillableId', 'A MembershipType', 
    'B MemberId', 'B BillableId', 'B MembershipType']);
  for(my $row = 0; $row < scalar(@{$conflicts}); $row++) {
    write_record($conflictWorksheet, $row + 1, [
      $conflicts->[$row]{'familyId'},
      $conflicts->[$row]{'a-memberId'},
      $conflicts->[$row]{'a-billableId'},
      $conflicts->[$row]{'a-membershipType'},
      $conflicts->[$row]{'b-memberId'},
      $conflicts->[$row]{'b-billableId'},
      $conflicts->[$row]{'b-membershipType'},
    ]);
  }
}

NOFAMILY: {
  my $noFamilyWorkbook = make_workbook('no_family');
  my $noFamilyWorksheet = make_worksheet($noFamilyWorkbook, ['MemberId']);
  for(my $row = 0; $row < scalar(@{$noFamily}); $row++) {
    write_record($noFamilyWorksheet, $row + 1, [
      $noFamily->[$row]{'MemberId'},
    ]);
  }
}

($members, $headers) = openMembersFile();

my $row = 1;
while(my $line = <$members>) {
  chomp $line;

  next unless ($line =~ /F193819563/);
  
  $csv->parse($line) || die "Line could not be parsed: $line";

  my $values = map_values($headers, [$csv->fields()]);
  # print Dumper($values, $families->{$values->{'FamilyId'}});exit;

  my $family = $families->{$values->{'FamilyId'}};

  $values->{'PrimaryAddress1'} = 'NOT AVAILABLE';
  $values->{'PrimaryAddress2'} = '';
  $values->{'PrimaryCity'} = '';
  $values->{'PrimaryState'} = '';
  $values->{'PrimaryZip'} = '';
  $values->{'PrimaryCountry'} = '';
  #$values->{'PrimaryEmail'} = '';

  if ($values->{'MemberId'} eq $family->{'primaryId'}) {
    $values->{'PrimaryAddress1'} = $family->{'address'}{'address1'};
    $values->{'PrimaryAddress2'} = $family->{'address'}{'address2'};
    $values->{'PrimaryCity'} = $family->{'address'}{'city'};
    $values->{'PrimaryState'} = $family->{'address'}{'state'};
    $values->{'PrimaryZip'} = $family->{'address'}{'zip'};
    $values->{'PrimaryZip'} = 'USA';
    #$values->{'PrimaryEmail'} = $family->{'email'};
  }

  my $record = make_record($values, \@allColumns, $columnMap);
  write_record($worksheet, $row++, $record);

}

close($members);

sub openMembersFile {
  open(my $members, '<:encoding(UTF-8)', 'data/AllMembers.csv')
    or die "Couldn't open data/AllMembers.csv: $!";
  my $headerLine = <$members>;
  $csv->parse($headerLine) || die "Line could not be parsed: $headerLine";
  my @headers = $csv->fields();

  return $members, \@headers;
}

sub billableMember {
  my $values = shift;

  my @nonBillableTypes = (
    'Non-Member',
  );

  unless (grep { $_ eq $values->{'MembershipType'} } @nonBillableTypes) {
    return $values->{'BillableMemberId'} eq $values->{'MemberId'};
  }

  return 0;
}
