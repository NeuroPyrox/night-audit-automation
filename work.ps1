$inspect = $null;

Function Inspect-Strings {
    Param ([string[]]$strings);
    $strings | % {Write-Host "$($_.Length) $_"}
}

Function Wait {
    Sleep -Milliseconds 200;
    Add-Type -AssemblyName System.Windows.Forms;
    $x = [System.Windows.Forms.Cursor]::Position.X;
    $y = [System.Windows.Forms.Cursor]::Position.Y;
    if (($x -eq 0) -or ($x -eq 1439) -or ($y -eq 0) -or ($y -eq 899)) {
        throw "Safeguard: stop program when mouse moves to edge of screen";
    }
}

Add-Type -AssemblyName System.Windows.Forms
Function Send-Keys {
	Param ([string]$keys);
    [System.Windows.Forms.SendKeys]::SendWait($keys)
    Wait;
}

$Mouse=@' 
[DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@;
$SendMouseClick = Add-Type -memberDefinition $Mouse -name "Win32MouseEventNew" -namespace Win32Functions -passThru;

Function Move-Mouse {
    Param ([int]$x, [int]$y);
    Process {
        Add-Type -AssemblyName System.Windows.Forms;
        $screen = [System.Windows.Forms.SystemInformation]::VirtualScreen;
        $null = $screen | Get-Member -MemberType Property;
        $screen.Width = $x;
        $screen.Height = $y;
        [Windows.Forms.Cursor]::Position = "$($screen.Width),$($screen.Height)";
        Wait;
    }
}
Function Down-Mouse {
    $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
    Wait;
}
Function Up-Mouse {
    $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
    Wait;
}
Function Left-Click {
    $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
    $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0);
    Wait;
}
function Right-Click {
    $SendMouseClick::mouse_event(0x00000008, 0, 0, 0, 0);
    $SendMouseClick::mouse_event(0x00000010, 0, 0, 0, 0);
    Wait;
}

Function Navigate-To-Room-Number {
	Param ([int]$roomNumber);
	Send-Keys ($roomNumber.ToString());
	Send-Keys "~";
	Send-Keys "~";
	Move-Mouse 710 250;
	Down-Mouse;
	Move-Mouse 1310 250;
	Up-Mouse;
	Right-Click;
	Move-Mouse 1250 260;
	Left-Click;
	$found = Get-Clipboard;
	if ($found -eq "NO MATCHES!                         ") {
		Send-Keys "{F4}";
		return $false;
	}
	if ($found.Substring(0, 3) -ne $roomNumber.ToString()) {
		throw "Expected to find a checked-in room of the same value";
	}
	Send-Keys "~";
	$true;
}

Function Copy-Housekeeping-Screen {
	Move-Mouse 10 385;
	Down-Mouse;
	Move-Mouse 1240 660;
	Up-Mouse;
	Right-Click;
	Move-Mouse 1250 400;
	Left-Click;
    $result = (Get-Clipboard) -split "`n";
    $result = $result -split "`n";
	if ($result[1].Substring(1, 12) -ne "Service Date") {
		throw "Not on the housekeeping screen";
	}
	$result;
}

Function Add-Rfsh {
    Param ([int]$tidyIndex, [int]$rfshIndex);
	for ($i = 0; i -lt $tidyIndex; i++) {
		Send-Keys "{DOWN}";
	}
	Send-Keys "{F2}";
	# TODO do we have a weekly dialog here?
	for ($i = 0; i -lt $tidyIndex; i++) {
		Send-Keys "{UP}";
	}
	for ($i = 0; i -lt $rfshIndex; i++) {
		Send-Keys "{DOWN}";
	}
	Send-Keys "{F2}";
	# TODO handle weekly dialog
	for ($i = 0; i -lt $rfshIndex; i++) {
		Send-Keys "{UP}";
	}
}

Function Add-Housekeeping {
    Param ([int]$daysCount);
	if ($daysCount -eq 0) {
	    return;
	}
	# TODO Add-Tidy-Service
	if ($daysCount -lt 3) {
	    return;
	}
	# TODO Add-Rfsh-Service
	$housekeeping = Copy-Housekeeping-Screen;
	$services = Parse-Services $housekeeping;
	# TODO use correct IndexOf substitute for generic lists
	$tidyIndex = [array]::IndexOf($services, "TIDY");
	$rfshIndex = [array]::IndexOf($services, "RFSH");
	# TODO navigate to first rfsh
	Add-Rfsh $tidyIndex $rfshIndex;
	if ($daysCount -lt 7) {
	    return;
	}
	# TODO navigate to second rfsh
	Add-Rfsh $tidyIndex $rfshIndex;
}

Function Array-Some {
    Param ([object[]]$array, $predicate);
    foreach ($item in $array) {
        if (&$predicate $item) {
            return $true;
        }
    }
    $false;
}

Function Trim-End {
    Param ($array, $isEnd);
    $result = New-Object System.Collections.Generic.List[System.Object];
    $reachedEnd = $false;
    foreach ($item in $array) {
        if (&$isEnd $item) {
			$reachedEnd = $true;
		} elseif ($reachedEnd) {
			throw "Didn't expect to find anything after the end";
		} else {
			$result.add($item);
		}
    }
    $result;
}

Function Parse-Services {
	Param ([string[]]$housekeeping);
	# TODO find the right index instead of 3
	$services = 5..9 | % {
		if ($housekeeping[$_].Length -ne 72) {
			throw "Expected 72 characters";
		}
		$housekeeping[$_].Substring(1, 9);
	}
	$services = Trim-End $services {Param($x); $x -eq "         "};
	if ($services.Count -ne ($services | Select-Object -Unique).Count) {
		throw "Should only have one of each type of available service";
	}
	$services;
}

Function Parse-Days-Count {
	Param ([string[]]$housekeeping);
	$days = $housekeeping[0];
	if ($days.Length -ne 72) {
		throw "Expected 72 characters";
	}
	$days = 0..8 | % {$days.Substring(21 + (6 * $_), 3)};
	$days = Trim-End $days {Param($x); $x -eq "   "};
	if (Array-Some $days {Param([string]$x); !($x -in "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")}) {
		throw "Unexpected value for a day";
	}
	$days.Count;
}

# Assumes services aren't weird
Function Parse-Schedule {
	Param ([string[]]$housekeeping);
	[string[][]] $schedule = (0..8 | % {
		$dayIndex = $_;
		5..9 |
			% {$housekeeping[$_].Substring(20 + (6 * $dayIndex), 4)} |
			where {$_ -ne "    "} |
			% {
				if (!($_ -in @("C/O ", "TIDY", "RFSH", "1XWE"))) {
					throw "Unrecognized service";
				}
				$_;
			}
	})
	Trim-End $schedule {Param([string[]]$x); $x.Count -eq 0};
}

Function Are-Services-Weird {
	Param ([object[]]$services);
	if (Array-Some $services {Param([string]$x); !($x -in "CHECK OUT", "TIDY     ", "RFSH     ", "1XWE     ")}) {
		return $true;
	}
	$foundCheckout = "CHECK OUT" -in $services;
	$foundTidy = "TIDY     " -in $services;
	$foundRfsh = "RFSH     " -in $services;
	$found1xwe = "1XWE     " -in $services;
	!($foundCheckout -and ( `
				(!$foundTidy -and !$foundRfsh -and !$found1xwe) `
				-or (!$foundTidy -and !$foundRfsh -and $found1xwe) `
				-or ($foundTidy -and !$foundRfsh -and !$found1xwe) `
				-or ($foundTidy -and $foundRfsh -and !$found1xwe) `
				-or ($foundTidy -and $foundRfsh -and $found1xwe) `
				));
}

Function Is-Checkout-Weird {
	Param ([string[][]]$schedule);
	if (Array-Some ($schedule | Select-Object -SkipLast 1) {Param($x); "C/O " -in $x}) {
		return $true;
	}
	if ($schedule.Count -eq 9) {
		return ($schedule[-1].Count -ne 1) -and ("C/O " -in $schedule[-1]);
	}
	return ($schedule[-1].Count -ne 1) -or !("C/O " -in $schedule[-1]);
}

# Assumes there won't be any unrecognized services
Function Are-Non-Checkouts-Weird {
	Param ([string[][]]$schedule);
	# If there's no matching schedule
	!(Array-Some @( `
				@(0, 0, 1, 0, 0, 0, 2, 0, 0), `
				@(2, 0, 0, 1, 0, 0, 0, 2, 0), `
				@(0, 2, 0, 0, 1, 0, 0, 0, 2), `
				@(0, 0, 2, 0, 0, 1, 0, 0, 0), `
				@(0, 0, 0, 2, 0, 0, 1, 0, 0), `
				@(1, 0, 0, 0, 2, 0, 0, 1, 0), `
				@(0, 1, 0, 0, 0, 2, 0, 0, 1), `
				@(0, 0, 0, 0, 0, 0, 3, 0, 0), `
				@(3, 0, 0, 0, 0, 0, 0, 3, 0), `
				@(0, 3, 0, 0, 0, 0, 0, 0, 3), `
				@(0, 0, 3, 0, 0, 0, 0, 0, 0), `
				@(0, 0, 0, 3, 0, 0, 0, 0, 0), `
				@(0, 0, 0, 0, 3, 0, 0, 0, 0), `
				@(0, 0, 0, 0, 0, 3, 0, 0, 0) `
				) {
            Param($weeklyPattern);
			# If there's no mismatching day
			for ($dayIndex = 0; $dayIndex -lt $schedule.Count; $dayIndex++) {
                $foundTidy = "TIDY" -in $schedule[$dayIndex];
			    $foundRfsh = "RFSH" -in $schedule[$dayIndex];
			    $found1xwe = "1XWE" -in $schedule[$dayIndex];
			    # If the day doesn't match
			    if (!( `
				        (($weeklyPattern[$dayIndex] -eq 0) -and $foundTidy -and !$foundRfsh -and !$found1xwe) `
					    -or (($weeklyPattern[$dayIndex] -eq 1) -and !$foundTidy -and $foundRfsh -and !$found1xwe) `
				        -or (($weeklyPattern[$dayIndex] -eq 2) -and !$foundTidy -and $foundRfsh -and $found1xwe) `
					    -or (($weeklyPattern[$dayIndex] -eq 3) -and !$foundTidy -and !$foundRfsh -and $found1xwe) `
				     )) {
				    return $false;
			    }
			}
			$true;
	})
}

Function Is-Schedule-Empty {
	Param ([string[][]]$schedule);
	if (("TIDY" -in $schedule[0]) -or ("RFSH" -in $schedule[0])) {
	    return $false;
	}
	$schedule | % {
		if (("TIDY" -in $_) -or ("RFSH" -in $_)) {
		    throw "Expected whole schedule to be empty if the first day was empty";
		}
	}
	$true;
}

Function Add-Housekeeping-If-None {
	Param ([int]$roomNumber);
	$housekeeping = Copy-Housekeeping-Screen;
	# TODO validate $housekeeping
	$services = Parse-Services $housekeeping;
	if (Are-Services-Weird $services) {
		return Write-Host "$roomNumber weird services";
	}
	$daysCount = Parse-Days-Count $housekeeping;
	$schedule = Parse-Schedule $housekeeping;
	if ($daysCount -ne $schedule.Count) {
		throw "Expected days and schedule to be the same length";
	}
	if (Is-Checkout-Weird $schedule) {
		# TODO make into an assertion if possible
		return Write-Host "$roomNumber weird C/O";
	}
	if ("C/O " -in $schedule[-1]) {
		$schedule = $schedule | Select-Object -SkipLast 1;
	}
	if (Are-Non-Checkouts-Weird $schedule) {
		return Write-Host "$roomNumber weird TIDY, RFSH, or 1XWE";
	}
	if (!(Is-Schedule-Empty $schedule)) {
		return Write-Host "$roomNumber normal"
	}
	if (("TIDY" -in $services) -or ("RFSH" -in $services)) {
	    return Write-Host "$roomNumber weird unused services"
	}
	# TODO first make sure everything above works as expected
	# Add-Housekeeping $daysCount
}

Function Main {
    $roomNumbers = $(101..103; 105; 126..129; 201..214; 216..229; 231; 301..329; 331; 401..429; 431);
    Move-Mouse 10 385;
    Left-Click;
    foreach ($roomNumber in $roomNumbers) {
        # TODO start at room #
	    $foundRoom = Navigate-To-Room-Number $roomNumber;
	    if ($foundRoom) {
		    Send-Keys "g";
		    Add-Housekeeping-If-None $roomNumber;
		    Send-Keys "{F4}";
            Wait;
            Wait;
            Wait;
            Wait;
		    Send-Keys "{F4}";
	    } else {
            Write-Host "$roomNumber not found"
        }
    }
}
