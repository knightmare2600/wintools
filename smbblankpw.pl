#!/usr/bin/perl

use constant NUM_TO_MAKE => 1;
my @addrs;
my $made = 0;
OUTER:
for my $i ( 1 .. 254 ) {
    for my $j ( 1 .. 254 ) {
        for my $k ( 1 .. 254 ) {
            for my $l ( 1 .. 254 ) {
                push @addrs, "$i.$j.$k.$l";
                  my $returncode = `smbclient -N -L "@addrs"`;
                last OUTER if ++$made >= NUM_TO_MAKE;
            }
        }
    }
}
