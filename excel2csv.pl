#!/usr/bin/env perl

use strict;
use warnings;
use Carp;
use Getopt::Long;
use Pod::Usage;
use File::Basename qw/fileparse/;
use File::Spec qw/rel2abs/;
use Spreadsheet::ParseExcel;
use Spreadsheet::XLSX;

my %args = ();
my $help = undef;
GetOptions(
           \%args,
           'excel=s',
           'sheet=s',
           'man|help' => \$help,
          ) or die pod2usage(1);

pod2usage(1) if $help;
pod2usage(
          -verbose => 2,
          -output  => \*STDERR
         ) unless defined $args{excel} || defined $args{sheet};

if (_getSuffix($args{excel}) eq ".xls") {
    my $file = File::Spec->rel2abs($args{excel});

    if (-e $file) {
        print _XLS(
                   file  => $file,
                   sheet => $args{sheet}
                  );
    } else {
        die "Error: Can not find file:$!";
    }
}
elsif (_getSuffix($args{excel}) eq ".xlsx") {
    my $file = File::Spec->rel2abs($args{excel});

    if (-e $file) {
        print _XLSX(
                    file  => $file,
                    sheet => $args{sheet}
                   );
    } else {
        die "Error: Can not find file:$!";
    }
}

sub _XLS {
    my %opts = (
                file  => undef,
                sheet => undef,
                @_,
               );

    my $aggregated = ();
    my $parser     = Spreadsheet::ParseExcel->new();
    my $workbook   = $parser->parse($opts{file});

    if (!defined $workbook) {
        croak "Error: $parser->error()";
    }

    foreach my $worksheet ($workbook->worksheet($opts{sheet})) {
        my ($row_min, $row_max) = $worksheet->row_range();
        my ($col_min, $col_max) = $worksheet->col_range();

        foreach my $row ($row_min .. $row_max){
            foreach my $col ($col_min .. $col_max){
                my $cell = $worksheet->get_cell($row, $col);
                if ($cell) {
                    $aggregated .= $cell->value().',';
                } else {
                    $aggregated .= ',';
                }
            }
            $aggregated .= "\n";
        }
    }
    return $aggregated;
}

sub _XLSX {
    my %opts = (
                file  => undef,
                sheet => undef,
                @_,
               );

    my $aggregated_x = ();
    my $excel = Spreadsheet::XLSX->new($opts{file});

    foreach my $sheet (@{ $excel->{Worksheet} }) {
        if ($sheet->{Name} eq $opts{sheet}) {
            $sheet->{MaxRow} ||= $sheet->{MinRow};

            foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
                $sheet->{MaxCol} ||= $sheet->{MinCol};
                foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {
                    my $cell = $sheet->{Cells}->[$row]->[$col];
                    if ($cell) {
                        $aggregated_x .= $cell->{Val}.',';
                    }
                }
                $aggregated_x .= "\n";
            }
        }
    }
    return $aggregated_x;
}

sub _getSuffix {
    my $f = shift;
    my ($basename, $dirname, $ext) = fileparse($f, qr/\.[^\.]*$/);
    return $ext;
}


__END__

=head1 NAME

xls2csv - Converting XLS/XLSX file to CSV

=head1 SYNOPSIS

perl xls2csv --excel data.xls|.xlsx --sheet Sheet1

=head1 OPTIONS

 -e,  --excel     Given a .xls or .xlsx file.       [Required]
 -s,  --sheet     Given a sheet name of the file.   [Required]
 -h,  --help      Show help messages.

=head1 DESCRIPTION

This program converts .xls/.xlsx file to csv,
automatically converting by the given file suffix.

=cut
