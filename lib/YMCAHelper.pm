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
  dump
  dd
  open_data_file
  process_data_file
  process_customer_file
  clean_customer
  get_gender
  is_member
  billable_member
  member_order_master_fields
  program_order_master_fields
  branch_name_map
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
  dump
  dd
  open_data_file
  process_data_file
  process_customer_file
  clean_customer
  is_member
  billable_member
  member_order_master_fields
  program_order_master_fields
  branch_name_map
);

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });
  
my $dbh = DBI->connect('dbi:SQLite:dbname=db/ymca.db','','');

my @customerCompanies;
process_data_file(
  'data/CustomerCompanies.csv',
  sub {
    my $values = shift;
    push(@customerCompanies, $values->{'TRX_ID'});
  }
);

#print Dumper(\@customerCompanies);exit;

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

  # print "$allColumns->[3]\n";exit;
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

  die "Couldn't map $tId" unless ($pId);

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
      return if (grep { $values->{'MemberId'} eq $_ } @customerCompanies);

      $func->($values);
    },
    'customers'
  );
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
    'BTW Community Center' => 'BT',
    'Middletown' => 'MD',
    'Atrium' => 'AT',
    'Fairfield Family' => 'FF',
    'Fitton Family' => 'FT',
    'East Butler' => 'EB',
    'Hamilton Central' => 'HC',
    'Metropolitan' => 'HC',
  };
}

sub member_order_master_fields {

  my @orderMasterFields = qw(
    OrderNo
    OrderDate
    OrgId
    OrgUnitId
    PerBillableMemberId
    BillAddressTypeCode
    ShipCustomerId
    OrderMethodCode
    OrderStatusCode
    OrderStatusDate
    OrdstsReasonCode
    ClOrderMethodCode
    CouponCode
    Application
    AckLetterMethodCode
    PoNumber
    ConfirmationNo
    AckLetterPrintDate
    ConfirmationDate
    OrderCompleteFlag
    AdvContractId
    AdvAgencyCustId
    BillSalesTerritory
    FndGiveEmployerCreditFlag
    ShipSalesTerritory
    PosFlag
    PosCountryCode
    PosState
    PosPostalCode
    AdvRateCardYearCode
    AdvAgencySubCustId
    EmployerCustomerId
    OldOrderNo
    MembershipType
    PaymentMethod 
    RenewalFee
    BranchCode
    MembershipBranch
    CompanyName
    NextBillDate
    JoinDate
    FamilyId
    PerMemberId
    SponsorDiscount
  );

  return @orderMasterFields;
}

sub program_order_master_fields {

  my @orderMasterFields = qw(
    OrderNo
    OrderDate
    OrgId
    OrgUnitId
    BillCustomerId
    BillAddressTypeCode
    ShipCustomerId
    OrderMethodCode
    OrderStatusCode
    OrderStatusDate
    OrdstsReasonCode
    ClOrderMethodCode
    CouponCode
    Application
    AckLetterMethodCode
    PoNumber
    ConfirmationNo
    AckLetterPrintDate
    ConfirmationDate
    OrderCompleteFlag
    AdvContractId
    AdvAgencyCustId
    BillSalesTerritory
    FndGiveEmployerCreditFlag
    ShipSalesTerritory
    PosFlag
    PosCountryCode
    PosState
    PosPostalCode
    AdvRateCardYearCode
    AdvAgencySubCustId
    EmployerCustomerId
    OldOrderNo
    Session
    ProgramEndDate
    LastName
    ItemDescription
    MemberId
    ReceiptNumber
    FeePaid
    DatePaid
    GlAccount
    ProgramStartDate
    BillableLastName
    BillableFirstName
    Branch
    BranchName
    Cycle
    ProgramDescription
    FirstName
    BillableMemberId
    OrderNo
    OrderDate
    StatusDate
    PerMemberId
    PerBillableMemberId
  );

  return @orderMasterFields;
}

1;