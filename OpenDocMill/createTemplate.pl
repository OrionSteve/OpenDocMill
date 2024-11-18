#!/usr/bin/env perl
use warnings;

use strict;
use JSON;
use IPC::Open3;
use Carp;
use Data::Dumper;
use Orion::Cmdline;

my $templateData = {
    fields => [ [ "field1", "Field One" ], [ "field2", "Field Two" ], ],
    tables => [ [ "table1", [ [ "f1", "Fld 1" ], [ "f2", "Fld 2" ], [ "f3", "Fld 3" ], ] ], ]
};

sub createTemplate {
    my ( $templateTemplateFilename, $data, $outTemplateFilename ) = @_;
    my $json = JSON::to_json($data);
    my ( $READ, $WRITE );
    my @cmd
        = ( "/usr/local/bin/python-orion", "runTemplateCreator.py", $templateTemplateFilename, $outTemplateFilename );
    my $pid = open3( $WRITE, $READ, 0, @cmd );
    print $WRITE $json;
    close $WRITE;
    waitpid $pid, 0;
    my $status = $?;
    my $output = join "", <$READ>;
    $output =~ s/^\s+//;
    $output =~ s/\s+$//g;
    return $output if $status == 0;
    confess "Command failed: data=" . Dumper($data) . "; cmd=[@cmd]; output=[$output]\n";
}

my $output = createTemplate( "templateTemplate.odt", $templateData, "outTemplate.odt" );
print "Command succeeded: output=[$output]\n";
