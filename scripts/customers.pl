#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::CSV_XS;
use Date::Manip;
use Text::Table;
use Term::ProgressBar;

my $cusIndTemplateName = 'DCT_CUS_INDIVIDUAL';

my $cusIndColumnMap = {
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

my @cusIndAllColumns = get_template_columns($cusIndTemplateName);

my $cusIndWorkbook = make_workbook($cusIndTemplateName);
my $cusIndWorksheet = make_worksheet($cusIndWorkbook, \@cusIndAllColumns);

my $cusRelTemplateName = 'DCT_CUS_RELATIONSHIP-20681';

my $cusRelColumnMap = {
  'RELATED_TRX_ID'            => { 'type' => 'record', 'source' => 'PrimaryId' },
  'TRX_ID'                    => { 'type' => 'record', 'source' => 'MemberId' },
  'RELATED_NAME'              => { 'type' => 'record', 'source' => 'PrimaryName' },
  'RELATIONSHIP_TYPE'         => { 'type' => 'static', 'source' => 'FAMILY' },
  'RELATIONSHIP_CODE'         => { 'type' => 'static', 'source' => 'FAMILY_MEMBER' },
  'RECIPROCAL_CODE'           => { 'type' => 'static', 'source' => 'FAMILY_MEMBER' },
  'BEGIN_DATE'                => { 'type' => 'static', 'source' => '1/1/2018' },
  'PRIMARY_CONTACT_FLAG'      => { 'type' => 'static', 'source' => 'N' },
  'PRIMARY_EMPLOYER_FLAG'     => { 'type' => 'static', 'source' => 'N' },
  'CL_AFFILIATE_MANAGER_FLAG' => { 'type' => 'static', 'source' => 'Y' },
};

my @cusRelAllColumns = get_template_columns($cusRelTemplateName);

my $cusRelWorkbook = make_workbook($cusRelTemplateName);
my $cusRelWorksheet = make_worksheet($cusRelWorkbook, \@cusRelAllColumns);

my $addrLnkTemplateName = 'DCT_ADDRESS_LINKING-43751';

my $addrLnkColumnMap = {
  'MASTER_TRX_ID'                        => { 'type' => 'record', 'source' => 'MemberId' },
  'LABEL_NAME'                           => { 'type' => 'record', 'source' => 'PrimaryName' },
  'ADDRESS_TYPE_CODE'                    => { 'type' => 'static', 'source' => 'HOME' },
  'ADDRESS_STATUS_CODE'                  => { 'type' => 'static', 'source' => 'GOOD' },
  'ADDRESS_STATUS_CHANGE_DATE'           => { 'type' => 'record', 'source' => 'CurrentDate' },
  'LINK_FROM_TRX_ID'                     => { 'type' => 'record', 'source' => 'MemberId' },
  'LINK_FROM_ADDRESS_TYPE'               => { 'type' => 'static', 'source' => 'HOME' },
  'PRIMARY_FLAG'                         => { 'type' => 'static', 'source' => 'Y' },
  'ONE_TIME_USE_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'CONFIDENTIAL_FLAG'                    => { 'type' => 'static', 'source' => 'N' },
  'SHIP_TO_FLAG'                         => { 'type' => 'static', 'source' => 'Y' },
  'BILL_TO_FLAG'                         => { 'type' => 'static', 'source' => 'Y' },
  'WEB_MOBILE_DIRECTORY_FLAG'            => { 'type' => 'static', 'source' => 'N' },
  'INCLUDE_IN_WEB_MOBILE_DIRECTORY_FLAG' => { 'type' => 'static', 'source' => 'N' },
  'DIRECTORY_PRIORITY'                   => { 'type' => 'static', 'source' => '0' },
  'RECUR_FLAG'                           => { 'type' => 'static', 'source' => 'N' },
  'AP_FLAG'                              => { 'type' => 'static', 'source' => 'N' },
  'PRIMARY_SEARCH_GROUP_OVERRIDE_FLAG'   => { 'type' => 'static', 'source' => 'N' },
};

my @addrLnkAllColumns = get_template_columns($addrLnkTemplateName);

my $addrLnkWorkbook = make_workbook($addrLnkTemplateName);
my $addrLnkWorksheet = make_worksheet($addrLnkWorkbook, \@addrLnkAllColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

my($membersFile, $headers, $totalRows) = openMembersFile();

print "Processing customers\n";
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });

my $members = {};
my $families = {};
my $familyOldestMember = {};
my $conflicts = [];
my $noFamily = [];
my $count = 1;
while(my $rowIn = $csv->getline($membersFile)) {

  $progress->update($count++);

  my $values = clean_customer(map_values($headers, $rowIn));
  # next unless ($values->{'FamilyId'} eq 'F365293034');
  # print Dumper($values); exit;

  $members->{$values->{'MemberId'}} = $values;

  next unless (isMember($values));
  
  # Camp members who do not need to be loaded
  unless ($values->{'FamilyId'}) {
    push(@{$noFamily}, {
      'MemberId' => $values->{'MemberId'},
    });
    next;
  }

  $values->{'FormalName'} = $values->{'FirstName'} . ' ' . $values->{'LastName'};

  addToFamilies($values, $families, $conflicts, $familyOldestMember);
}

close($membersFile);

# Put non-members into families
print "Processing non-members\n";
$progress = Term::ProgressBar->new({ 'count' => scalar(keys %{$members}) });
$count = 1;
my $overSubscribedFamilies = {};
foreach my $memberId (keys %{$members}) {
  $progress->update($count);
  my $member = $members->{$memberId};
  next if (isMember($member));

  # if the family exists, assign the first membership type
  # and primary ID to this non-member
  if (exists($families->{$member->{'FamilyId'}})) {
    my @familyTypes = keys $families->{$member->{'FamilyId'}};
    
    $member->{'OverSubscribed'} = 0;
    if (@familyTypes > 1) {
      $overSubscribedFamilies->{$member->{'FamilyId'}} =
        $families->{$member->{'FamilyId'}};
      $member->{'OverSubscribed'} = 1;
      next;
    }

    $member->{'MembershipType'} = $familyTypes[0];
    $member->{'BillableMemberId'} 
      = $families->{$member->{'FamilyId'}}{$member->{'MembershipType'}}{'primaryId'};
  }

  addToFamilies($member, $families, $conflicts, $familyOldestMember);
}

# Find families with no primary and use oldest member
print "Processing primary family members\n";
$progress = Term::ProgressBar->new({ 'count' => scalar(keys %{$families}) });
$count = 1;

my $primaryByBillable = 0;
my $primaryByBirth = 0;
my $missingPrimary = 0;
my $familyCount = 0;
foreach my $familyId (keys %{$families}) {
  foreach my $membershipType (keys %{$families->{$familyId}}) {
    $progress->update($count);
    $familyCount++;

    my $family = $families->{$familyId}{$membershipType};

    if ($family->{'primaryId'}) {
      $primaryByBillable++;
    } else {
      if (exists($familyOldestMember->{$familyId})) {
        $family->{'primaryId'}
          = $familyOldestMember->{$familyId}{'memberId'};
        $primaryByBirth++;
      }
    }

    $missingPrimary++ unless ($family->{'primaryId'});
  }
}

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

OVERSUBSCRIBED: {
  my $maxTypes = 0;
  foreach my $familyId (keys %{$overSubscribedFamilies}) {
    $maxTypes = (keys %{$overSubscribedFamilies->{$familyId}})
      if ((keys %{$overSubscribedFamilies->{$familyId}}) > $maxTypes);
  }

  my $columns = ['FamilyId'];
  for my $i (1..$maxTypes) {
    push(@{$columns}, "PrimaryId $i", "Membership Type $i");
  }

  my $overSubWorkbook = make_workbook('over_subscribed');
  my $overSubWorksheet = make_worksheet($overSubWorkbook, $columns);

  my $row = 1;
  foreach my $familyId (keys %{$overSubscribedFamilies}) {
    my $values = [
      $familyId,
    ];

    foreach my $type (keys %{$overSubscribedFamilies->{$familyId}}) {
      push(
        @{$values}, 
        $overSubscribedFamilies->{$familyId}{$type}{'primaryId'},
        $type
      );
    }

    write_record($overSubWorksheet, $row++, $values);
  }
}

my $currentDate = UnixDate(ParseDate('today'), '%Y-%m-%d');

print "Generating customer files\n";
$progress = Term::ProgressBar->new({ 'count' => scalar(keys %{$members}) });
$count = 1;
my $indRow = 1;
my $lnkRow = 1;
foreach my $memberId (keys %{$members}) {
  $progress->update($count);
  my $member = $members->{$memberId};

  # next unless ($member->{'FamilyId'} eq 'F365293034');
  # print Dumper($member, $families->{$member->{'FamilyId'}});exit;

  next if ($member->{'OverSubscribed'});
  
  my $family = $families->{$member->{'FamilyId'}}{uc $member->{'MembershipType'}};
  my $primaryMember = $members->{$family->{'primaryId'}};
  my $isPrimary = $member->{'MemberId'} eq $family->{'primaryId'};

  $member->{'CurrentDate'} = $currentDate;

  # These Primary fields are only to filled in with primary member values
  $member->{'PrimaryId'} = $family->{'primaryId'};
  $member->{'PrimaryAddress1'} = 'NOT AVAILABLE';
  $member->{'PrimaryAddress2'} = '';
  $member->{'PrimaryCity'} = '';
  $member->{'PrimaryState'} = '';
  $member->{'PrimaryZip'} = '';
  $member->{'PrimaryCountry'} = '';
  $member->{'PrimaryName'} = $primaryMember->{'FormalName'};
  
  # Move the email address to the primary member unless the primary has one
  $member->{'Email'} = $primaryMember->{'Email'} unless ($member->{'Email'});

  if ($isPrimary) {
    $member->{'PrimaryAddress1'} = $member->{'Address1'};
    $member->{'PrimaryAddress2'} = $member->{'Address2'};
    $member->{'PrimaryCity'} = $member->{'City'};
    $member->{'PrimaryState'} = $member->{'State'};
    $member->{'PrimaryZip'} = $member->{'Zip'};
    $member->{'PrimaryCountry'} = 'USA';
    $member->{'PrimaryName'} = $member->{'FormalName'};
  }

  my $cusIndRecord = make_record($member, \@cusIndAllColumns, $cusIndColumnMap);
  write_record($cusIndWorksheet, $indRow++, $cusIndRecord);

  unless ($isPrimary) {
    my $cusRelRecord = make_record($member, \@cusRelAllColumns, $cusRelColumnMap);
    write_record($cusRelWorksheet, $lnkRow, $cusRelRecord);

    my $addrLnkRecord = make_record($member, \@addrLnkAllColumns, $addrLnkColumnMap);
    write_record($addrLnkWorksheet, $lnkRow, $addrLnkRecord);

    $lnkRow++;
  }
}

close($membersFile);

sub openMembersFile {
  my $csv = Text::CSV_XS->new ({ auto_diag => 1 });
  
  # Subtract one for the heading row
  my $totalRows = `cat data/AllMembers.csv | wc -l` - 1;

  open(my $members, '<:encoding(UTF-8)', 'data/AllMembers.csv')
    or die "Couldn't open data/AllMembers.csv: $!";
  my $headers = $csv->getline($members);

  return $members, $headers, $totalRows;
}

sub clean_customer {
  my $values = shift;

  # Trim all name fields
  # Clear all address fields if address1 is empty
  # Remove trailing - in zip
  # Remove non digits in phone
  # Discard on 10 digit phones
  # Ensure valid email format
  
  return $values;
}

sub isMember {
  my $values = shift;

  return 0 if ($values->{'MembershipType'} eq 'Non-Member');
  return 0 if ($values->{'MembershipType'} =~ /program/i);

  return 1;
}

sub billableMember {
  my $values = shift;

  return 0 unless isMember($values);

  return $values->{'BillableMemberId'} eq $values->{'MemberId'};
}

sub addToFamilies {
  my $values = shift;
  my $families = shift;
  my $conflicts = shift;
  my $familyOldestMember = shift;

  my $familyId = $values->{'FamilyId'};
  my $membershipType = uc $values->{'MembershipType'};

  unless (exists($families->{$familyId})) {
    $families->{$familyId} = {};    
  };

  unless (exists($families->{$familyId}{$membershipType})) {
    $families->{$familyId}{$membershipType} = {
      'primaryId' => '',
      'members' => [],
    };
  }

  my $family = $families->{$familyId}{$membershipType};

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

  push(@{$family->{'members'}}, $values->{'MemberId'});
  
  if (billableMember($values)) {
    if ($family->{'primaryId'} && 
        $family->{'primaryId'} ne $values->{'MemberId'}) {
      # compare($values, $members->{$family->{'primaryId'}});
      # exit;
      my $conflictedMember = $members->{$family->{'primaryId'}};
      push(@{$conflicts}, {
        'familyId' => $familyId,
        'a-memberId' => $values->{'MemberId'},
        'a-billableId' => $values->{'BillableMemberId'},
        'a-membershipType' => $values->{'MembershipType'},
        'b-memberId' => $conflictedMember->{'MemberId'},
        'b-billableId' => $conflictedMember->{'BillableMemberId'},
        'b-membershipType' => $conflictedMember->{'MembershipType'},
      });
      return;
    } 

    $family->{'primaryId'} = $values->{'MemberId'};
  }
}
