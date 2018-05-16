#!/usr/bin/perl -w

use strict;

use Data::Dumper;
use Spreadsheet::XLSX;

foreach my $type (qw(templates data)) {
  opendir(TYPEDIR, "excel-$type") || die "Couldn't open excel-$type: $!";
  my @files = grep !/^\./, readdir(TYPEDIR);
  closedir(TYPEDIR);
  
  foreach my $fileInName (@files) {
    print "Converting excel-$type/$fileInName\n";
    my $workbook = Spreadsheet::XLSX->new("excel-$type/$fileInName");
    my $sheet = $workbook->{'Worksheet'}[0];

    (my $fileOutName = $fileInName) =~ s/xls(x|m)$/txt/;
    open(my $fileOut, '>', "$type/$fileOutName") or die "Couldn't open $type/$fileOutName: $!";
    foreach my $row ($sheet->{'MinRow'} .. $sheet->{'MaxRow'}) {
      my @rowValues;
      foreach my $col ($sheet->{'MinCol'} .. $sheet->{'MaxCol'}) {
        push(@rowValues, $sheet->{'Cells'}[$row][$col]->{'Val'});
      }
      print $fileOut join("\t", @rowValues) . "\n";
    }
    close($fileOut);
  }

}
