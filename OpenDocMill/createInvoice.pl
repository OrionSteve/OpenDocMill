#!/usr/bin/env perl
use warnings;

use strict;
use JSON;
use IPC::Open3;
use Carp;
use Data::Dumper;

my $invoiceData = [
    {   name   => "INVOICE",    # Page name; must match the HEADING1 in the ODT
        fields => {
            invoiceNo => "SL0010040",
            account   => "508",
            date      => "08/04/09",
            comment   => "Note: Items supplied under pre-agreed terms and conditions",
            vatRate   => "15%",
            taxable   => "360.25",
            zeroRated => "0",
            vat       => "54.04",
            terms     => "Payment due 28 days from date of invoice",
            total     => "414.29",
        },
        tables => {
            items => [
                {   product => "Beef -Aberdeen Angus -5kg",
                    date    => "10 Apr",
                    price   => "111.51",
                    req     => "3215783",
                    ref     => "RAYD-2009-0406",
                    qty     => "1"
                },
                {   product => "Apples Cox -each",
                    date    => "10 Apr",
                    price   => "6.78",
                    req     => "3215784",
                    ref     => "RAYD-2009-0406",
                    qty     => "100"
                },
                {   product => "Apples Bramley -per Kilo",
                    date    => "10 Apr",
                    price   => "21.56",
                    req     => "3215784",
                    ref     => "RAYD-2009-0406",
                    qty     => "10"
                },
            ],
            address => [
                { line => "CHURCHILL ENGINEERING" },
                { line => "63, TUDOR CLOSE" },
                { line => "LONDON" },
                { line => "NW3 4AG" },
                { line => "UNITED KINGDOM" },
            ]
        },
    },
    { name => "#footer" },
];

sub writeReport {
    my ( $templateFilename, $reportData, $outFilename ) = @_;
    my $json = JSON::to_json($reportData);
    my ( $READ, $WRITE );
    my @cmd = ( "/usr/local/bin/python-orion", "runOpenDocMill.py", $templateFilename, $outFilename );
    my $pid = open3( $WRITE, $READ, 0, @cmd );
    print $WRITE $json;
    close $WRITE;
    waitpid $pid, 0;
    my $status = $?;
    my $output = join "", <$READ>;
    $output =~ s/^\s+//;
    $output =~ s/\s+$//g;
    return $output if $status == 0;
    confess "Command failed: data=" . Dumper($reportData) . "; cmd=[@cmd]; output=[$output]\n";
}

my $output = writeReport( "invoiceTemplate.odt", $invoiceData, "out.odt" );
print "Command succeeded: output=[$output]\n";
