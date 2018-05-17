package YMCAHelper;
use strict;
use warnings;
use Exporter;
use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;

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

1;