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

my $cusIndTemplateName = 'DCT_CUS_INDIVIDUAL';

my $cusIndColumnMap = {
  'ORG_ID'                               => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'ORG_UNIT_ID'                          => { 'type' => 'static', 'source' => 'GMVYMCA' },
  'TRX_ID'                               => { 'type' => 'record', 'source' => 'MemberId' },
  'CUSTOMER_ID'                          => { 'type' => 'record', 'source' => 'PerMemberId' },
  'NAME_PREFIX'                          => { 'type' => 'record', 'source' => 'Prefix' },
  'FIRST_NAME'                           => { 'type' => 'record', 'source' => 'FirstName' },
  'LAST_NAME'                            => { 'type' => 'record', 'source' => 'LastName' },
  'NAME_SUFFIX'                          => { 'type' => 'record', 'source' => 'Suffix' },
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
  'ADDRESS_TYPE_CODE'                    => { 'type' => 'record', 'source' => 'PrimaryAddressTypeCode' },
  'ADDRESS_STATUS_CODE'                  => { 'type' => 'record', 'source' => 'PrimaryAddressStatusCode' },
  'COMPANY_NAME'                         => { 'type' => 'record', 'source' => 'Corporation' },
  'PHONE_AREA_CODE'                      => { 'type' => 'record', 'source' => 'HomePhoneAreaCode' },
  'PRIMARY_PHONE'                        => { 'type' => 'record', 'source' => 'HomePhoneNumber' },
  'PRIMARY_PHONE_LOCATION_CODE'          => { 'type' => 'record', 'source' => 'PhoneLocationCode' },
  'PRIMARY_EMAIL_ADDRESS'                => { 'type' => 'record', 'source' => 'Email' },
  'PRIMARY_EMAIL_LOCATION_CODE'          => { 'type' => 'record', 'source' => 'EmailLocationCode' },
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
  'PRIMARY_MOBILE_PHONE'                 => { 'type' => 'record', 'source' => 'CellPhone' },
  'PRIMARY_MOBILE_PHONE_LOCATION_CODE'   => { 'type' => 'record', 'source' => 'CellLocationCode' },
  'PUBLISH_PRIMARY_MOBILE_PHONE_FLAG'    => { 'type' => 'static', 'source' => 'N' },
};

my @cusIndAllColumns = get_template_columns($cusIndTemplateName);

my $cusIndWorkbook = make_workbook($cusIndTemplateName);
my $cusIndWorksheet = make_worksheet($cusIndWorkbook, \@cusIndAllColumns);

my $cusRelTemplateName = 'DCT_CUS_RELATIONSHIP-20681';

my $cusRelColumnMap = {
  'RELATED_CUSTOMER_ID'       => { 'type' => 'record', 'source' => 'PerPrimaryId' },
  'RELATED_TRX_ID'            => { 'type' => 'record', 'source' => 'PrimaryId' },
  'CUSTOMER_ID'               => { 'type' => 'record', 'source' => 'PerMemberId' },
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
  'MASTER_CUSTOMER_ID'                   => { 'type' => 'record', 'source' => 'PerMemberId' },
  'MASTER_TRX_ID'                        => { 'type' => 'record', 'source' => 'MemberId' },
  'LABEL_NAME'                           => { 'type' => 'record', 'source' => 'FormalName' },
  'ADDRESS_TYPE_CODE'                    => { 'type' => 'static', 'source' => 'HOME' },
  'ADDRESS_STATUS_CODE'                  => { 'type' => 'static', 'source' => 'GOOD' },
  'ADDRESS_STATUS_CHANGE_DATE'           => { 'type' => 'record', 'source' => 'CurrentDate' },
  'LINK_FROM_CUSTOMER_ID'                => { 'type' => 'record', 'source' => 'PerPrimaryId' },
  'LINK_FROM_TRX_ID'                     => { 'type' => 'record', 'source' => 'PrimaryId' },
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

my($membersFile, $headers, $totalRows) = open_data_file('data/AllMembers.csv');

print "Processing customers\n";
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });

my $members = {};
my $families = {};
my $conflicts = [];
my $noFamily = [];
my $count = 1;
while(my $rowIn = $csv->getline($membersFile)) {

  $progress->update($count++);

  my $values = clean_customer(map_values($headers, $rowIn));
  # next unless ($values->{'FamilyId'} eq 'F152136702');
  # next unless ($values->{'TrxEmail'} eq 'lljennings99@gmail.com');
  # dump($values); exit;

  $members->{$values->{'MemberId'}} = $values;

  next unless (is_member($values));
  
  # Camp members who do not need to be loaded
  unless ($values->{'FamilyId'}) {
    push(@{$noFamily}, {
      'MemberId' => $values->{'MemberId'},
    });
    next;
  }

  # We will determine if the member is the primary later.
  $values->{'IsFamilyPrimary'} = 0;

  addToFamilies($values, $families, $conflicts);
}

close($membersFile);

# Put non-members into families
print "Processing non-members\n";
$progress = Term::ProgressBar->new({ 'count' => scalar(keys %{$members}) });
$count = 1;
my $overSubscribedFamilies = {};
foreach my $memberId (keys %{$members}) {
  $progress->update($count++);
  my $member = $members->{$memberId};
  next if (is_member($member));

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
      = $families->{$member->{'FamilyId'}}{$member->{'MembershipType'}}{'PrimaryId'};
  }

  addToFamilies($member, $families, $conflicts);
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
    $progress->update($count++);
    $familyCount++;

    my $family = $families->{$familyId}{$membershipType};

    if ($family->{'PrimaryId'}) {
      $primaryByBillable++;
    } else {
      my $familyMembers = [];
      foreach my $memberId (@{$family->{'Members'}}) {
        push(@{$familyMembers}, $members->{$memberId});
      }

      my $oldestMember = oldestMember($familyMembers);

      $family->{'PrimaryId'} = $oldestMember->{'MemberId'};
      $primaryByBirth++ if ($family->{'PrimaryId'});
    }

    die "Unable to find primary for $familyId" unless ($family->{'PrimaryId'});

    $members->{$family->{'PrimaryId'}}{'IsFamilyPrimary'} = 1;

    # Assign email to the primary if they don't have one,
    # but another family member does. In the case of changing
    # family email that goes into Personify, we will also
    # change the TRX email data and not worry about fixing it.
    unless ($members->{$family->{'PrimaryId'}}{'Email'}) {
      foreach my $familyMemberId (sort @{$family->{'Members'}}) {
        $members->{$family->{'PrimaryId'}}{'Email'}
          = $members->{$familyMemberId}{'Email'};
        $members->{$family->{'PrimaryId'}}{'EmailLocationCode'}
          = $members->{$familyMemberId}{'EmailLocationCode'};
        $members->{$family->{'PrimaryId'}}{'TrxEmail'}
          = $members->{$familyMemberId}{'TrxEmail'};
        last if ($members->{$family->{'PrimaryId'}}{'Email'});
      }
    }

    # Remove the email from any other family members. This will
    # stomp any any family member's own email, but there is no
    # way to be completely accurate. The choice is to prefer
    # the primary.
    if ($members->{$family->{'PrimaryId'}}{'Email'}) {
      foreach my $familyMemberId (@{$family->{'Members'}}) {
        next if ($familyMemberId eq $family->{'PrimaryId'});
        $members->{$familyMemberId}{'Email'} = '';
        $members->{$familyMemberId}{'TrxEmail'} = '';
        $members->{$familyMemberId}{'EmailLocationCode'} = '';
      }
    }
  }
}

CONFLICTS: {
  my $conflictWorkbook = make_workbook('conflicted_primary');
  my $conflictWorksheet = make_worksheet($conflictWorkbook, 
    ['FamilyId', 'A MemberId', 'A BillableId', 'A MembershipType', 
    'B MemberId', 'B BillableId', 'B MembershipType']);
  for(my $row = 0; $row < scalar(@{$conflicts}); $row++) {
    write_record($conflictWorksheet, $row + 1, [
      $conflicts->[$row]{'FamilyId'},
      $conflicts->[$row]{'A-MemberId'},
      $conflicts->[$row]{'A-BillableId'},
      $conflicts->[$row]{'A-MembershipType'},
      $conflicts->[$row]{'B-MemberId'},
      $conflicts->[$row]{'B-BillableId'},
      $conflicts->[$row]{'B-membershipType'},
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
        $overSubscribedFamilies->{$familyId}{$type}{'PrimaryId'},
        $type
      );
    }

    write_record($overSubWorksheet, $row++, $values);
  }
}

print "Checking for duplicate email\n";
my $trxEmail = {};
my $activeEmail = {};
my $memberIds = {};
$progress = Term::ProgressBar->new({ 'count' => scalar(keys %{$members}) });
$count = 1;
foreach my $memberId (keys %{$members}) {
  $progress->update($count++);

  my $member = $members->{$memberId};
  
  $memberIds->{$member->{'PerMemberId'}}++;
  if ($memberIds->{$member->{'PerMemberId'}} > 1) {
    print "Personify MemberId $member->{'PerMemberId'} ($memberId) created duplicate\n";
  }

  # Clear out any remaining non-member email address
  unless (is_member($member)) {
    $member->{'TrxEmail'} = '';
    $member->{'Email'} = '';
    $member->{'EmailLocationCode'} = '';
  }

  if ($member->{'TrxEmail'}) {
    my $email = $member->{'TrxEmail'};
    $trxEmail->{$email} = [] unless (exists($trxEmail->{$email}));
    push(@{$trxEmail->{$email}}, $member);
  }

  if ($member->{'Email'}) {
    my $email = $member->{'Email'};
    $activeEmail->{$email} = [] unless (exists($activeEmail->{$email}));
    push(@{$activeEmail->{$email}}, $member);
  }
}

DUPEMAIL: {
  my $dupEmailWorkbook = make_workbook('duplicate_email');
  my $dupEmailWorksheet = make_worksheet($dupEmailWorkbook, 
    ['Email', 'FamilyId', 'MemberId', 'Primary', 'Last', 'First', 'Type']);
  my $row = 1;
  foreach my $email (keys %{$trxEmail}) {
    next unless (scalar(@{$trxEmail->{$email}}) > 1);
    foreach my $member (@{$trxEmail->{$email}}) {
      write_record($dupEmailWorksheet, $row, [
        $email,
        $member->{'FamilyId'},
        $member->{'MemberId'},
        $member->{'IsFamilyPrimary'} ? 'primary' : '',
        $member->{'LastName'},
        $member->{'FirstName'},
        $member->{'MembershipType'},
      ]);
      $row++;
    }
  }

  foreach my $email (keys %{$activeEmail}) {
    next unless (scalar(@{$activeEmail->{$email}}) > 1);
    # Resolve duplicates here
    # find oldest of $activeEmail->{$email}
    my $oldestMember = oldestMember($activeEmail->{$email});

    # clear email from all others
    foreach my $member (@{$activeEmail->{$email}}) {
      next if ($member->{'MemberId'} eq $oldestMember->{'MemberId'});
      $member->{'Email'} = '';
      $member->{'EmailLocationCode'} = '';
    }
  }
}

my $currentDate = UnixDate(ParseDate('today'), '%Y-%m-%d');

print "Generating customer files\n";
$progress = Term::ProgressBar->new({ 'count' => scalar(keys %{$members}) });
$count = 1;
my $emailCheck = {};
my $indRow = 1;
my $lnkRow = 1;
foreach my $memberId (keys %{$members}) {
  $progress->update($count++);
  my $member = $members->{$memberId};

  # next unless ($member->{'FamilyId'} eq 'F161136482');
  # print Dumper($member, $families->{$member->{'FamilyId'}});exit;

  next if ($member->{'OverSubscribed'});

  my $isMember = is_member($member);
  
  if ($member->{'Email'}) {
    $emailCheck->{$member->{'Email'}}++;
    print "$member->{'Email'} still duplicated\n"
      if ($emailCheck->{$member->{'Email'}} > 1);
  }

  my $family = $families->{$member->{'FamilyId'}}{uc $member->{'MembershipType'}};
  my $primaryMember = $members->{$family->{'PrimaryId'}};
  my $isPrimary = $isMember && $member->{'MemberId'} eq $family->{'PrimaryId'};

  $member->{'CurrentDate'} = $currentDate;

  # These Primary fields are only to filled in with primary member values
  $member->{'PrimaryId'} = $primaryMember->{'MemberId'};
  $member->{'PerPrimaryId'} = $primaryMember->{'PerMemberId'};
  $member->{'PrimaryAddress1'} = 'NOT AVAILABLE';
  $member->{'PrimaryAddress2'} = '';
  $member->{'PrimaryCity'} = '';
  $member->{'PrimaryState'} = '';
  $member->{'PrimaryZip'} = '';
  $member->{'PrimaryCountry'} = '';
  $member->{'PrimaryAddressTypeCode'} = '';
  $member->{'PrimaryAddressStatusCode'} = '';

  $member->{'PrimaryName'} = $primaryMember->{'FormalName'};
  
  if ($isPrimary) {
    $member->{'PrimaryAddress1'} = $member->{'Address1'};
    $member->{'PrimaryAddress2'} = $member->{'Address2'};
    $member->{'PrimaryCity'} = $member->{'City'};
    $member->{'PrimaryState'} = $member->{'State'};
    $member->{'PrimaryZip'} = $member->{'Zip'};
    $member->{'PrimaryCountry'} = $member->{'Country'};
    $member->{'PrimaryName'} = $member->{'FormalName'};
    $member->{'PrimaryAddressTypeCode'} = $member->{'AddressTypeCode'};
    $member->{'PrimaryAddressStatusCode'} = $member->{'AddressStatusCode'};
  }

  if ($member->{'Email'} && !$member->{'EmailLocationCode'}) {
    dd($member);
  }

  my $cusIndRecord = make_record($member, \@cusIndAllColumns, $cusIndColumnMap);

  write_record($cusIndWorksheet, $indRow++, $cusIndRecord);

  if ($isMember && !$isPrimary) {
    my $cusRelRecord = make_record($member, \@cusRelAllColumns, $cusRelColumnMap);
    write_record($cusRelWorksheet, $lnkRow, $cusRelRecord);

    my $addrLnkRecord = make_record($member, \@addrLnkAllColumns, $addrLnkColumnMap);
    write_record($addrLnkWorksheet, $lnkRow, $addrLnkRecord);

    $lnkRow++;
  }
}

close($membersFile);

sub oldestMember {
  my $members = shift;

  my $oldestMember = {};

  foreach my $member (@{$members}) {
    # If the member has no birth day, don't process them
    unless ($member->{'DateOfBirth'}) {
      # But if there is no oldest member yet, use this one to start
      $oldestMember = $member unless ($oldestMember);
      next;
    }

    # If the oldest member doesn't have a birth day, use this one
    unless ($oldestMember->{'DateOfBirth'}) {
      $oldestMember = $member;
      next;
    }

    # Prefer real member over non-members
    if (!$oldestMember->{'IsMember'} && $member->{'IsMember'}) {
      $oldestMember = $member;
      next;
    }

    # At this point, both current and oldest should have birth days
    my $memberBirthDate = ParseDate($member->{'DateOfBirth'});
    my $oldestBirthDate = ParseDate($oldestMember->{'DateOfBirth'});

    # If the current member birth day is less than the oldest, use it
    if (Date_Cmp($memberBirthDate, $oldestBirthDate) == -1) {
      $oldestMember = $member;
    }
  }

  return $oldestMember
}

sub addToFamilies {
  my $values = shift;
  my $families = shift;
  my $conflicts = shift;

  my $familyId = $values->{'FamilyId'};
  my $membershipType = uc $values->{'MembershipType'};

  unless (exists($families->{$familyId})) {
    $families->{$familyId} = {};    
  };

  unless (exists($families->{$familyId}{$membershipType})) {
    $families->{$familyId}{$membershipType} = {
      'PrimaryId' => '',
      'Members' => [],
    };
  }

  my $family = $families->{$familyId}{$membershipType};

  push(@{$family->{'Members'}}, $values->{'MemberId'});
  
  if (billable_member($values)) {
    if ($family->{'PrimaryId'} && 
        $family->{'PrimaryId'} ne $values->{'MemberId'}) {
      # compare($values, $members->{$family->{'PrimaryId'}});
      # exit;
      my $conflictedMember = $members->{$family->{'PrimaryId'}};
      push(@{$conflicts}, {
        'FamilyId' => $familyId,
        'A-MemberId' => $values->{'MemberId'},
        'A-BillableId' => $values->{'BillableMemberId'},
        'A-MembershipType' => $values->{'MembershipType'},
        'B-MemberId' => $conflictedMember->{'MemberId'},
        'B-BillableId' => $conflictedMember->{'BillableMemberId'},
        'B-membershipType' => $conflictedMember->{'MembershipType'},
      });
      return;
    } 

    $family->{'PrimaryId'} = $values->{'MemberId'};
  }
}
