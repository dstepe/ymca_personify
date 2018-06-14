package YMCAHelper;
use strict;
use warnings;
use Exporter;
use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::Table;
use Text::CSV_XS;
use Term::ProgressBar;
use DBI;
use Date::Manip;

our @ISA= qw( Exporter );

# these CAN be exported.
our @EXPORT_OK = qw( 
  get_template_columns
  make_worksheet
  make_workbook
  write_record
  split_values
  map_values
  make_record
  lookup_id
  compare
  by_program_start_date
  dump
  dd
  open_data_file
  process_data_file
  process_customer_file
  clean_customer
  get_gender
  is_member
  billable_member
  member_order_fields
  program_order_fields
  donation_order_fields
  branch_name_map
  resolve_branch_name
  skip_program
  skip_cycle
  is_company
);

# these are exported by default.
our @EXPORT = qw( 
  get_template_columns
  make_workbook
  make_worksheet
  write_record
  split_values
  map_values
  make_record
  lookup_id
  compare
  by_program_start_date
  dump
  dd
  open_data_file
  process_data_file
  process_customer_file
  clean_customer
  is_member
  billable_member
  member_order_fields
  program_order_fields
  donation_order_fields
  branch_name_map
  resolve_branch_name
  skip_program
  skip_cycle
  is_company
);

my $csv = Text::CSV_XS->new ({ binary => 1 });
  
my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

sub get_template_columns {
  my $templateName = shift;
  
  my $rowNum = 1;

  return split(
    "\t", 
    (read_file('templates/' . $templateName . '.txt', 'chomp' => 1))[$rowNum]
  );
}

sub make_workbook {
  my $templateName = shift;

  my $workbook = Excel::Writer::XLSX->new('complete/' . $templateName . '.xlsx');

  return $workbook;
}

sub make_worksheet {
  my $workbook = shift;
  my $allColumns = shift;

  my $worksheet = $workbook->add_worksheet();

  my $format = $workbook->add_format();
  $format->set_bold();
  $format->set_color( 'red' );

  for(my $i = 0; $i < scalar(@{$allColumns}); $i++) {
    $worksheet->write(0, $i, $allColumns->[$i], $format);
  }

  return $worksheet;
}

sub write_record {
  my $worksheet = shift;
  my $row = shift;
  my $record = shift;

  for(my $i = 0; $i < scalar(@{$record}); $i++) {
    if (!defined($record->[$i])) {
      print "$i not defined\n";
    }
    if ($record->[$i] =~ /^0\d/) {
      $worksheet->write_string($row, $i, $record->[$i]);
    } else {
      $worksheet->write($row, $i, $record->[$i]);
    }   
  }
}

sub map_values {
  my $headers = shift;
  my $values = shift;

  my $mapped = {};

  die "header/value mismatch" unless (scalar(@{$headers}) == scalar(@{$values}));

  for(my $i = 0; $i < scalar(@{$headers}); $i++) {
    $mapped->{$headers->[$i]} = $values->[$i];
  }

  return $mapped;
}

sub split_values {
  my $row = shift;

  my @values = split("\t", $row);

  my $mapped = {};
  while (my $name = shift) {
    $mapped->{$name} = shift @values;
  }

  return $mapped;
}

sub make_record {
  my $values = shift;
  my $allColumns = shift;
  my $columnMap = shift;

  # print "$allColumns->[32]\n";exit;
  my @record;
  foreach my $field (@{$allColumns}) {
    unless (exists($columnMap->{$field})) {
      push(@record, '');
      next;
    }

    if ($columnMap->{$field}{'type'} eq 'record') {
      if (!exists($values->{$columnMap->{$field}{'source'}})) {
        print "values don't contain $columnMap->{$field}{'source'} for $field\n";
        print Dumper($values->{$columnMap->{$field}{'source'}});
        exit;
      }
      push(@record, $values->{$columnMap->{$field}{'source'}});
      next;
    }

    push(@record, $columnMap->{$field}{'source'});
  }

  return \@record;
}

sub lookup_id {
  my $tId = shift;

  $tId = sprintf('%09d', $tId) if ($tId =~ /^\d/);

  my($pId) = $dbh->selectrow_array(q{
    select p_id
      from ids
      where t_id = ?
    }, undef, $tId);

  print "Couldn't map $tId\n" unless ($pId);

  return $pId;
}

sub compare {
  my $a = shift;
  my $b = shift;
  my $aLabel = shift || 'A';
  my $bLabel = shift || 'B';

  my $table = Text::Table->new('Attribute', $aLabel, $bLabel);
  foreach my $key (sort keys %{$a}) {
    $table->add($key, $a->{$key}, $b->{$key});
  }

  print $table;
}

sub by_program_start_date {
  my $a = shift;
  my $b = shift;

  return 0 unless ($a->{'StartDateTime'} && $b->{'StartDateTime'});
  return 1 unless ($a->{'StartDateTime'});
  return -1 unless ($b->{'StartDateTime'});

  my $aStartDate = ParseDate($a->{'StartDateTime'});
  my $bStartDate = ParseDate($b->{'StartDateTime'});

  Date_Cmp($aStartDate, $bStartDate);
}

sub dump {
  my $obj = shift;

  my $table = Text::Table->new('Attribute', 'Value');
  foreach my $key (sort keys %{$obj}) {
    my $value = $obj->{$key};
    $value = join(', ', @{$value}) if (ref($value) eq 'ARRAY');
    $table->add($key, $value);
  }

  print $table;  
}

sub dd {
  my $obj = shift;

  &dump($obj);
  exit;

}

sub open_data_file {
  my $file = shift;
  my $headerMap = shift || {};

  # Subtract one for the heading row
  my $totalRows = `cat $file | wc -l` - 1;

  open(my $fileHndl, '<:encoding(UTF-8)', $file)
    or die "Couldn't open $file: $!";
  
  my $headers = $csv->getline($fileHndl);

  for (my $i = 0; $i < scalar(@{$headers}); $i++) {
    $headers->[$i] = $headerMap->{$headers->[$i]} 
      if (exists($headerMap->{$headers->[$i]}));
  }

  return $fileHndl, $headers, $totalRows;
}

sub process_data_file {
  my $file = shift;
  my $func = shift;
  my $heading = shift || $file;
  my $headerMap = shift || {};

  my($dataFile, $headers, $totalRows) = open_data_file($file, $headerMap);

  my $showProgress = $totalRows > 100;

  my $progress;
  if ($showProgress) {
    print "Processing $heading\n";
    $progress = Term::ProgressBar->new({ 'count' => $totalRows });
  }

  my $count = 1;
  while(my $rowIn = $csv->getline($dataFile)) {

    $progress->update($count++) if ($showProgress);

    my $values = map_values($headers, $rowIn);
    
    $func->($values); 
  }

  close($dataFile);
}

sub process_customer_file {
  my $func = shift;

  process_data_file(
    'data/AllMembers.csv',
    sub {
      my $values = clean_customer(shift);
      
      # Skip companies in as customers
      return if (is_company($values->{'MemberId'}));

      $func->($values);
    },
    'customers'
  );
}

sub is_company {
  my $trxId = shift;

  my($count) = $dbh->selectrow_array(q{
    select count(*)
      from companies
      where t_id = ?
    }, undef, $trxId);

  return $count;
}

sub clean_customer {
  my $values = shift;

  # Trim all name fields
  foreach my $trimField (qw(Prefix FirstName LastName Suffix CasualName 
    Address1 Address2 City State Zip)) {
    $values->{$trimField} =~ s/^\s+//;
    $values->{$trimField} =~ s/\s+$//;
  }

  # Remove trailing '.'
  foreach my $trimField (qw(Prefix Suffix Address1 Address2)) {
    $values->{$trimField} =~ s/\.$//;
  }

  # Clear all address fields if address1 is empty
  $values->{'AddressTypeCode'} = 'HOME';
  $values->{'AddressStatusCode'} = 'GOOD';
  $values->{'Country'} = 'USA';
  unless ($values->{'Address1'}) {
    $values->{'Address1'} = 'NOT AVAILABLE';
    $values->{'Address2'} = '';
    $values->{'City'} = '';
    $values->{'State'} = '';
    $values->{'Zip'} = '';
    $values->{'Country'} = '';
    $values->{'AddressTypeCode'} = '';
    $values->{'AddressStatusCode'} = '';
  }

  # Remove +4 and trailing - in zip
  $values->{'Zip'} =~ s/-.*$//;

  # Remove non digits in phone
  # Discard non 10 digit phones
  foreach my $phoneField (qw(EmergencyPhone CellPhone WorkPhone HomePhone)) {
    $values->{$phoneField} =~ s/[^\d]//g;
    $values->{$phoneField} = '' unless ($values->{$phoneField} =~ /\d{10}/);
  }
  
  # Split home phone into area code and number
  $values->{'HomePhoneAreaCode'} = '';
  $values->{'HomePhoneNumber'} = '';
  if ($values->{'HomePhone'} =~ /(\d{3})(\d{7})/) {
    $values->{'HomePhoneAreaCode'} = $1;
    $values->{'HomePhoneNumber'} = $2;
  }

  # Ensure valid email format (very basic)
  $values->{'Email'} = ''
    unless ($values->{'Email'} =~ /.*\@.*\..*/);
  $values->{'Email'} =~ tr/[A-Z]/[a-z]/;
  
  # Remove non-member email to reduce duplicates, but keep
  # it for reporting to Y staff
  $values->{'TrxEmail'} = $values->{'Email'};
  # $values->{'Email'} = '' unless (is_member($values));

  # Add location code indicators if present
  $values->{'PhoneLocationCode'} = $values->{'HomePhoneNumber'} ? 'HOME' : '';
  $values->{'EmailLocationCode'} = $values->{'Email'} ? 'HOME' : '';
  $values->{'CellLocationCode'} = $values->{'CellPhone'} ? 'HOME' : '';
  
  $values->{'Gender'} = get_gender($values->{'Gender'});

  $values->{'FormalName'} = join(' ', $values->{'FirstName'}, $values->{'LastName'});

  # We may manipulate non-member membership types for family associations,
  # but must keep the real membership type for other purposes.
  $values->{'IsMember'} = is_member($values);

  # Convert TRX IDs to Personify IDs
  $values->{'PerMemberId'} = lookup_id($values->{'MemberId'});

  # dump($values);exit;
  return $values;
}

sub get_gender {
  my $code = shift;

  return 'MALE' if ($code eq 'M');
  return 'FEMALE' if ($code eq 'F');
  return 'OTHER';
}

sub is_member {
  my $values = shift;

  return 0 if (lc $values->{'MembershipType'} eq lc 'Non-Member');
  return 0 if ($values->{'MembershipType'} =~ /program/i);

  return 1;
}

sub billable_member {
  my $values = shift;

  return 0 unless is_member($values);

  return $values->{'BillableMemberId'} eq $values->{'MemberId'};
}

sub branch_name_map {
  return {
    'Camp Campbell' => 'CG',
    'BTW Community Center' => 'BT',
    'Middletown' => 'MD',
    'Atrium' => 'AT',
    'Fairfield Family' => 'FF',
    'Fitton Family' => 'FT',
    'East Butler' => 'EB',
    'Hamilton Central' => 'HC',
    'Metropolitan' => 'AS',
  };
}

sub resolve_branch_name {
  my $values = shift;
  my $from = shift || 'ItemDescription';

  my $name = '';

  $name = 'Camp Campbell' if ($values->{$from} =~ /Campbell/);
  $name = 'BTW Community Center' if ($values->{$from} =~ /Booker/);
  $name = 'Atrium' if ($values->{$from} =~ /Atrium/);
  $name = 'Fairfield Family' if ($values->{$from} =~ /Fairfield/);
  $name = 'Fitton Family' if ($values->{$from} =~ /Fitton/);
  $name = 'Metropolitan' if ($values->{$from} =~ /Association/);
  $name = 'Hamilton Central' if ($values->{$from} =~ /Central/);
  $name = 'East Butler' if ($values->{$from} =~ /East/);
  $name = 'Middletown' if ($values->{$from} =~ /Middletown/);

  return $name;
}

sub common_order_fields {
  return qw(
    OrderDate
    OrderNo
    PerBillableMemberId
    PerMemberId
    StatusDate
  );
}

sub member_order_fields {

  my @orderFields = common_order_fields();

  push(@orderFields, qw(
    MembershipTypeDes
    PaymentMethod
    RenewMembershipFee
    BranchCode
    MembershipBranch
    CompanyName
    NextBillDate
    JoinDate
    FamilyId
    SponsorDiscount
  ));

  return @orderFields;
}

sub program_order_fields {

  my @orderFields = common_order_fields();

  push(@orderFields, qw(
    Session
    ProgramStartDate
    ProgramEndDate
    ReceiptNumber
    ProgramFee
    FeePaid
    Balance
    DatePaid
    ItemDescription
    Cycle
    ProductCode
  ));

  return @orderFields;
}

sub donation_order_fields {

  my @orderFields = common_order_fields();

  push(@orderFields, qw(
    DonorName
    CampaignBalance
    CampaignPledge
    CampaignPledgeStatus
    ItemDescription
    PledgeType
    PledgeTypeFrequency
    PledgeNextBillDate
    ReceiptNumber
    FeePaid
    Balance
    DatePaid
    ProductCode
    CampaignCode
    FundCode
    Comments
  ));

  return @orderFields;
}

sub skip_cycle {
  my $cycle = shift;

  return 1 if ($cycle =~ /^2018 Guest Passes/);
  return 1 if ($cycle =~ /^2018 Personal Training/);
  return 1 if ($cycle =~ /^2018 Private Swim Lessons/);
  return 1 if ($cycle =~ /^2018 Spring/);
  return 1 if ($cycle =~ /^2018 Spring Soccer/);
  return 1 if ($cycle =~ /^2018 Winter 2/);
  return 1 if ($cycle =~ /^2018 Winter I/);
  return 1 if ($cycle =~ /^2018 Youth Basketball/);

  return 0;
}

sub skip_program {
  my $program = shift;

  foreach my $skip (
    q{17th Annual Turkey Trot 5K race},
    q{2 Hour Cycle},
    q{Adult Adaptive Aquatics},
    q{Adult Soccer League},
    q{Adult Swim Lessons- 4 week special},
    q{Adventure Night},
    q{Adventures In Storytelling},
    q{Aquatic Easter Egg Hunt},
    q{AQUATICS FOR PEOPLE WITH DISABILITIES},
    q{Art Camp},
    q{ASHI CPR},
    q{Atrium Camp School Out Days},
    q{BADMINTON - RECREATIONAL OPEN},
    q{Baseball/Softball Camp},
    q{Breakfast With Santa},
    q{Bump, Set, Splash},
    q{Car Show Trophy Sponsorship},
    q{Cardio Gold (Seniors)},
    q{Chair Yoga},
    q{Chefs in Training},
    q{Childrens Cntr Sports All Sorts},
    q{Connect. Condition. Challenge. Compete},
    q{Couch to 5k with Breanna},
    q{Daniel Plan},
    q{Dive-in movie night},
    q{Dodgeball},
    q{Double Decade Cycling},
    q{Downtown Showdown Teen Basketball},
    q{Easter Bunny Lunch},
    q{Eel},
    q{Enhance Fitness},
    q{Express Strength Circuit},
    q{Fairfield Family Y Walking Club},
    q{Fall Fest},
    q{Fall Weight Loss Challenge Sept. 18th-Nov.10th},
    q{Family game night/ special events},
    q{Floating Pumpkin Patch},
    q{Frog Jog},
    q{Good Friday Prayer Breakfast},
    q{Group CORE/CORE Focus Together},
    q{Group Power/Strength Train Together},
    q{Guppy},
    q{Gym Essentials},
    q{Gymnastics - JR HIGH SCHOOL},
    q{Gymnastics - PRIVATE LESSONS},
    q{Harvest Bash},
    q{Healthy Nutrition Class},
    q{Homeschool Gym and Swim},
    q{Indoor 5K- New Year's Day},
    q{Indoor Triathlon},
    q{Injury Screens},
    q{Kick Boxing},
    q{Kickball},
    q{Lifeguard},
    q{LIVESTRONG Bootcamp},
    q{Livestrong Car Show},
    q{Locker Rental},
    q{Luau Party},
    q{Lunch & Learn},
    q{Lunch N Learns},
    q{Miler'+s Club},
    q{Minnow},
    q{Nutrition Lecture Series},
    q{Open Dodge Ball},
    q{Operation Splash},
    q{Outside Groups / Swim},
    q{P.H.I.T.},
    q{Pajama and Movie Night},
    q{pee wee mini arts camp},
    q{Preteen Night/Kid's Night Out},
    q{Private Spanish Lessons},
    q{Private Sports Instruction},
    q{Private Swim},
    q{pumpkin painting},
    q{Pump},
    q{Runners Club},
    q{Seniors Focus Group},
    q{Small Group Training},
    q{Snacks With Santa},
    q{Speed and Agility Training},
    q{Splash and Dash Indoor Triathlon},
    q{Spring Fling Dance A Thon},
    q{Step},
    q{Stranded Island Camp},
    q{Stroke Development},
    q{Summer 1,000 Rep Challenge},
    q{Summer Wellness Challenge},
    q{Triathlon Clinic},
    q{TRX Small Group},
    q{Twilight Easter Egg Hunt},
    q{Two Hour Cycle},
    q{Underwater Easter Egg Hunt},
    q{Water in Motion},
    q{Water Volleyball},
    q{Y Swim Lesson Instructor Certification},
    q{Youth Stage 2 (Polliwog 2)},
  ) {
    return 1 if ($program =~ /^$skip/i);
  }

  return 0;
}

1;