package YMCAHelper;
use strict;
use warnings;
use Exporter;
use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Text::Table;

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
  convert_id
  compare
  dump
  open_members_file
  clean_customer
  get_gender
  is_member
  billable_member
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
  convert_id
  compare
  dump
  open_members_file
  clean_customer
  is_member
  billable_member
);

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
    $worksheet->write($row, $i, $record->[$i]);
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

  my @record;
  foreach my $field (@{$allColumns}) {
    unless (exists($columnMap->{$field})) {
      push(@record, '');
      next;
    }

    if ($columnMap->{$field}{'type'} eq 'record') {
      push(@record, $values->{$columnMap->{$field}{'source'}});
      next;
    }

    push(@record, $columnMap->{$field}{'source'});
  }

  return \@record;
}

sub convert_id {
  my $id = shift;
  return $id;

  # $id =~ s/^P/4/;

  return sprintf('%012d', $id);
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

sub open_members_file {
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
  $values->{'PerMemberId'} = convert_id($values->{'MemberId'});

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

1;