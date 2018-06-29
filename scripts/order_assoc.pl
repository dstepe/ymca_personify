#!/usr/bin/perl -w

use strict;

use lib 'lib';

use YMCAHelper;

use File::Slurp;
use Data::Dumper;
use Excel::Writer::XLSX;
use Date::Manip;
use Text::CSV_XS;
use Term::ProgressBar;

my $templateName = 'DCT_ORDER_MBR_ASSOCIATE-41924';

my $columnMap = {
  'ORDER_NO'                           => { 'type' => 'record', 'source' => 'OrderNo' },
  'ORDER_LINE_NO'                      => { 'type' => 'static', 'source' => '1' },
  'ASSOCIATE_CUSTOMER_ID'              => { 'type' => 'record', 'source' => 'PerMemberId' },
  'ASSOCIATE_CLASS_CODE'               => { 'type' => 'static', 'source' => 'FAMILY' },
};

my @allColumns = get_template_columns($templateName);

my $workbook = make_workbook($templateName);
my $worksheet = make_worksheet($workbook, \@allColumns);

my $csv = Text::CSV_XS->new ({ auto_diag => 1 });

# Load assoc orders file
my $assocOrder = {};
my($assocOrdersFile, $headers, $totalRows) = open_data_file('data/assoc_orders.csv');
while(my $rowIn = $csv->getline($assocOrdersFile)) {
  my $values = map_values($headers, $rowIn);

  my $membershipType = uc $values->{'MembershipType'};

  unless (exists($assocOrder->{$values->{'BillingMemberId'}}{$values->{'FamilyId'}}{$membershipType})) {
    $assocOrder->{$values->{'BillingMemberId'}}{$values->{'FamilyId'}}{$membershipType} = [];
  }

  push(@{$assocOrder->{$values->{'BillingMemberId'}}{$values->{'FamilyId'}}{$membershipType}}, $values->{'MemberId'});

}
close($assocOrdersFile);

# For each member order, find any assocs and add them here
my $ordersFile;
($ordersFile, $headers, $totalRows) = open_data_file('data/member_orders.csv');

my $familyOrders = {};
my $progress = Term::ProgressBar->new({ 'count' => $totalRows });
my $row = 1;
my $count = 1;
my $skipped = {};
while(my $rowIn = $csv->getline($ordersFile)) {

  $progress->update($count++);

  my $values = map_values($headers, $rowIn);

  my $membershipType = uc $values->{'MembershipTypeDes'};

  unless (exists($assocOrder->{$values->{'PerBillableMemberId'}})) {
    $skipped->{'Billiable not found'}++;
    next;
  }
  unless (exists($assocOrder->{$values->{'PerBillableMemberId'}}{$values->{'FamilyId'}})) {
    $skipped->{'Family not found'}++;
    next;
  }
  unless (exists($assocOrder->{$values->{'PerBillableMemberId'}}{$values->{'FamilyId'}}{$membershipType})) {
    $skipped->{'Membership not found'}{$membershipType}++;
    next;
  }

  my $assocMembers = $assocOrder->{$values->{'PerBillableMemberId'}}{$values->{'FamilyId'}}{$membershipType};

  foreach my $assocMember (@{$assocMembers}) {
    # Primary members are part of the main order and not added here
    next if ($assocMember eq $values->{'PerBillableMemberId'});

    my $record = {
      'OrderNo' => $values->{'OrderNo'},
      'PerMemberId' => $assocMember,
    };

    write_record(
      $worksheet,
      $row++,
      make_record($record, \@allColumns, $columnMap)
    );
  }

}
close($ordersFile);
print Dumper($skipped) if (%{$skipped});

# $row = 1;
# process_customer_file(
#   sub {
#     my $values = shift;

#     my $familyId = $values->{'FamilyId'};

#     dump($values);
#     return unless (exists($familyOrders->{$familyId}));
#     return if ($values->{'PerMemberId'} eq $familyOrders->{$familyId}{'BillingId'});
    
#     print Dumper($familyOrders->{$familyId});
#     print "customer $values->{'PerMemberId'} $familyId\n";
#     exit;

#     $values->{'OrderNo'} = $familyOrders->{$familyId}{'OrderNo'};

#     write_record(
#       $worksheet,
#       $row++,
#       make_record($values, \@allColumns, $columnMap)
#     );
#   }
# );
