$inspect = $null;

Function Skip-Last {
    Param ([object[]]$array);
    if ($array.Count -eq 0) {
        throw "Can't remove from an empty array"
    }
    if ($array.Count -eq 1) {
        return Write-Output -NoEnumerate @();
    }
	return Write-Output -NoEnumerate $array[-$array.Count..-2];
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
    $lastIndex = $array.Count - 1;
    while (&$isEnd $array[$lastIndex]) {
        $lastIndex--;
    }
    $result = [System.Collections.ArrayList]@();
    for ($i = 0; $i -le $lastIndex; $i++) {
        $null = $result.Add($array[$i]);
    }
    return Write-Output -NoEnumerate $result;
}

Function Assert-Valid-Request-Codes {
    Param([string[]]$requestCodes);
    $requestCodes | % {
        if (!($_ -in (( `
            "A0,A1,A5,A9,B4,B5,C1,D4,D7,D9,E1,E6,F2,G3,H1,H2,I1,I2,I4,J8,J9,K1,K2,K8,L2,L3" `
            + ",M1,M5,M8,MK,N1,N2,N3,N4,O9,P6,P8,R1,R3,R4,S5,S7,U2,V9,X1,X2,X4,X5,Y1,Y2,ZQ" `
        ) -split ","))) {
            throw "Unexpected request code: $_";
        }
    }
}

Function Parse-First-6-Requests {
	Param ([string]$raw);
    if ($raw.Length -ne 23) {
        throw "Expected 23 characters!";
    }
    $result = 0..5 | % {
        return $raw.Substring((4 * $_) + 1, 2);
    }
    $result = Trim-End $result {
        Param([string]$x);
        return $x -eq "  ";
    }
    Assert-Valid-Request-Codes $result;
    if ("  " -in $result) {
        throw "Unexpected space between requests!"
    }
    return $result;
}

Function Parse-F3-Requests {
	Param ([string]$raw);
    $withSpaces = 0..8 | % {
        return $raw.Substring(80 * $_, 2);
    };
    $withUnderscore = Trim-End $withSpaces {
        Param([string]$x);
        return $x -eq "  ";
    };
    $result = if ($withUnderscore[-1] -eq "__") {
        Skip-Last $withUnderscore;
    } else {
        $withUnderscore;
    }
    Assert-Valid-Request-Codes $result;
    return $result;
}

Function Parse-Services {
	Param ([string[]]$housekeeping);
	$services = 8..12 | % {
		if ($housekeeping[$_].Length -ne 72) {
			throw "Expected 72 characters";
		}
		$housekeeping[$_].Substring(1, 17);
	}
	$services = Trim-End $services {
        Param([string]$x);
        return $x -eq "                 ";
    };
	if ($services.Count -ne ($services | Select-Object -Unique).Count) {
		throw "Should only have one of each type of available service";
	}
	return Write-Output -NoEnumerate $services;
}

Function Parse-Days-Count {
	Param ([string[]]$housekeeping);
	$days = $housekeeping[3];
	if ($days.Length -ne 72) {
		throw "Expected 72 characters";
	}
	$days = 0..8 | % {
        return $days.Substring(21 + (6 * $_), 3);
    };
	$days = Trim-End $days {
        Param($x);
        return $x -eq "   ";
    };
	if (Array-Some $days {
        Param([string]$x);
        return !($x -in "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");
    }) {
		throw "Unexpected value for a day";
	}
	return $days.Count;
}

# Assumes services aren't weird
Function Parse-Unflattened-Schedule {
	Param ([string[]]$housekeeping);
    $schedule = [System.Collections.ArrayList]@();
	0..8 | % {
		$dayIndex = $_;
        $dayServices = [System.Collections.ArrayList]@();
		8..12 |
			% {$housekeeping[$_].Substring(20 + (6 * $dayIndex), 4)} |
			where {$_ -ne "    "} |
			% {
				if (!($_ -in @("C/O ", "TIDY", "RFSH", "1XWE"))) {
					throw "Unrecognized service";
				}
                $null = $dayServices.Add($_)
			};
        $null = $schedule.Add($dayServices);
	}
    $trimmed = Trim-End $schedule {
        Param([string[]]$x);
        return $x.Count -eq 0;
    };
    if ("C/O " -in $trimmed[-1]) {
	    return Write-Output -NoEnumerate $trimmed;
    }
	return Write-Output -NoEnumerate $schedule;
}

Function Parse-Schedule {
	Param ([string[]]$housekeeping);
    $schedule = Parse-Unflattened-Schedule $housekeeping;
    for ($i = 0; $i -lt $schedule.Count; $i++) {
        if ($schedule[$i].Count -eq 0) {
            $schedule[$i] = "    ";
        } elseif ($schedule[$i].Count -eq 1) {
            $schedule[$i] = $schedule[$i][0];
        } else {
            return "Overlapping services!";
        }
    }
    return Write-Output -NoEnumerate $schedule;
}

Function Are-Services-Weird {
	Param ([object[]]$services);
    $options = @( `
        "CHECK OUT        ", `
        "TIDY             ", `
        "HOUSEKEEPING REFR", `
        "1XWEEK           " `
    );
	if (Array-Some $services {Param([string]$x); !($x -in $options)}) {
		return $true;
	}
	$foundCheckout = "CHECK OUT        " -in $services;
	$foundTidy = "TIDY             " -in $services;
	$foundRfsh = "HOUSEKEEPING REFR" -in $services;
	$found1xwe = "1XWEEK           " -in $services;
	return !( `
        ($foundCheckout -and !(!$foundTidy -and $foundRfsh -and $found1xwe)) `
        -or (!$foundCheckout -and !$foundTidy -and !$foundRfsh -and $found1xwe) `
        -or (!$foundCheckout -and $foundTidy -and $foundRfsh -and $found1xwe) `
    );
}

# Assumes there won't be any unrecognized services
Function Are-Non-Checkouts-Weird {
	Param ([string[]]$schedule);
	# If there's no matching schedule
	!(Array-Some @( `
				@(1, 1, 2, 1, 1, 1, 3, 1, 1), `
				@(3, 1, 1, 2, 1, 1, 1, 3, 1), `
				@(1, 3, 1, 1, 2, 1, 1, 1, 3), `
				@(1, 1, 3, 1, 1, 2, 1, 1, 1), `
				@(1, 1, 1, 3, 1, 1, 2, 1, 1), `
				@(2, 1, 1, 1, 3, 1, 1, 2, 1), `
				@(1, 2, 1, 1, 1, 3, 1, 1, 2), `
				@(0, 0, 0, 0, 0, 0, 3, 0, 0) `
				) {
			# If there's no mismatching day
            Param($weeklyPattern);
			for ($dayIndex = 0; $dayIndex -lt $schedule.Count; $dayIndex++) {
			    # If the day doesn't match
                if ( `
                    $schedule[$dayIndex] `
                    -ne @("    ", "TIDY", "RFSH", "1XWE")[$weeklyPattern[$dayIndex]] `
                ) {
                    return $false;
                }
			}
			return $true;
	})
}

# Assumes non-checkouts aren't weird
Function Is-Schedule-Empty {
	Param ([string[]]$schedule);
    if ($schedule[0] -ne "    ") {
	    return $false;
    }
	$schedule | % {
		if (($_ -eq "TIDY") -or ($_ -eq "RFSH")) {
		    throw "Expected whole schedule to be empty if the first day was empty";
		}
	}
	$true;
}

Function Wait {
    Sleep -Milliseconds 100;
    Add-Type -AssemblyName System.Windows.Forms;
    $x = [System.Windows.Forms.Cursor]::Position.X;
    $y = [System.Windows.Forms.Cursor]::Position.Y;
    if (($x -eq 0) -or ($x -eq 1439) -or ($y -eq 0) -or ($y -eq 899)) {
        throw "Safeguard: stop program when mouse moves to edge of screen";
    }
}

$last10Keys = @("", "", "", "", "", "", "", "", "", "");
$last10KeysIndex = 0;
Add-Type -AssemblyName System.Windows.Forms
Function Send-Keys {
	Param ([string]$keys);
    $Global:last10Keys[$Global:last10KeysIndex] = $keys;
    $Global:last10KeysIndex = ($Global:last10KeysIndex + 1) % 10;
    [System.Windows.Forms.SendKeys]::SendWait($keys);
    Wait;
}
Function Send-Keys-Sequentially {
    Param ([string]$keys);
    ($keys -split ",") | % {Send-Keys $_;};
}
Function Unroll-Last-10-Keys {
    return Write-Output -NoEnumerate (0..9 | % {
        return $Global:last10Keys[($Global:last10KeysIndex + $_) % 10];
    });
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
Function Right-Click {
    $SendMouseClick::mouse_event(0x00000008, 0, 0, 0, 0);
    $SendMouseClick::mouse_event(0x00000010, 0, 0, 0, 0);
    Wait;
}

Function Retry-Get-Clipboard {
    $result = Get-Clipboard;
    if ($result.Length -ne 0) {
        return $result;
    }
    $result = Get-Clipboard;
    if ($result.Length -ne 0) {
        return $result;
    }
    $result = Get-Clipboard;
    if ($result.Length -ne 0) {
        return $result;
    }
    $result = Get-Clipboard;
    if ($result.Length -ne 0) {
        $Global:inspect = $result;
        throw "Please implement one more retry because it worked."
    }
    Wait;
    $result = (Get-Clipboard);
    if ($result.Length -ne 0) {
        $Global:inspect = $result;
        throw "Please implement one more retry because it worked."
    }
    throw "Expected a non-empty result";
}

Function Copy-From-Fosse {
    Param([int]$x1, [int]$x2, [int]$y1, [int]$y2, [int]$z1, [int]$z2);
	Move-Mouse $x1 $x2;
	Down-Mouse;
	Move-Mouse $y1 $y2;
	Up-Mouse;
	Right-Click;
    Wait;
	Move-Mouse $z1 $z2;
	Left-Click;
    return Retry-Get-Clipboard;
}

Function Navigate-To-Room-Number {
	Param ([int]$roomNumber);
	Send-Keys ($roomNumber.ToString());
	Send-Keys "~";
	Send-Keys "~";
	$found = Copy-From-Fosse 710 250 1310 520 1250 530;
    while ($found.Substring(363, 9) -eq "Room/Stay") {
        Send-Keys "{F4}{F4}";
	    Send-Keys ($roomNumber.ToString());
	    Send-Keys "~";
	    Send-Keys "~";
	    $found = Copy-From-Fosse 710 250 1310 520 1250 530;
    }
    $row0 = $found.Substring(0, 36);
    if ($row0.Substring(0, 3) -eq "Res") {
        Write-Host "Need to retry";
        throw "implement retrying";
    }
	if ($row0 -eq "NO MATCHES!                         ") {
		Send-Keys "{F4}";
		return $false;
	}
	if ($row0.Substring(0, 3) -eq $roomNumber.ToString()) {
	    Send-Keys "~";
	    return $true;
	}
    if ($row0.Substring(0, 3) -ne "C/O") {
        $Global:inspect = @($row0, $roomNumber);
        # Could be that $row0 is from the previous room
        Send-Keys "{F4}";
        # TODO check previous room number
	    throw "Expected to find a checked out room if the room number doesn't match";
    }
    $row1 = $found.Substring(80, 36);
    if ($row1 -eq "                                    ") {
		Send-Keys "{F4}";
		return $false;
    }
	if ($row1.Substring(0, 3) -eq $roomNumber.ToString()) {
	    Send-Keys "~";
	    return $true;
	}
    if ($row1.Substring(0, 3) -eq "C/O") {
	    throw "Unhandled case of 2 or more checkouts";
    }
	throw "Expected the same room number or `"C/O`"";
}

Function Has-J8 {
    $first6RequestsRaw = Copy-From-Fosse 660 500 1040 500 1050 510;
    # No profile was found
    if ($first6RequestsRaw -eq "wed by Acct Code)      ") {
        Send-Keys "~";
        $first6RequestsRaw = Copy-From-Fosse 660 500 1040 500 1050 510;
    }
    while ($first6RequestsRaw.Length -ne 23) {
        $first6RequestsRaw = Copy-From-Fosse 660 500 1040 500 1050 510;
    }
	$first6Requests = Parse-First-6-Requests $first6RequestsRaw;
    if ("J8" -in $first6Requests) {
        return $true;
    }
    if ($first6Requests.Count -lt 6) {
        return $false;
    }
    Send-Keys-Sequentially "E,pmont059,~,{UP},{UP},{UP},{F3}";
    $f3Requests = Parse-F3-Requests (Copy-From-Fosse 300 300 330 530 340 540);
    if ($f3Requests[0] -in $first6Requests) {
        Send-Keys "{F4}";
        return "J8" -in $f3Requests;
    }
    throw "Implement scrolling up";
}

Function Copy-Housekeeping-Screen {
    $clip = Copy-From-Fosse 270 300 1240 660 1250 400;
    $Global:inspect = $clip;
    if ($clip.GetType().name -ne "Object[]") {
        if ($clip.Substring(229, 9) -eq "Room/Stay") {
            return Copy-Housekeeping-Screen;
        } else {
		    throw "Not on the housekeeping screen";
        }
    }
	if ($clip[4].Substring(1, 12) -ne "Service Date") {
		throw "Not on the housekeeping screen";
	}
	return Write-Output -NoEnumerate $clip;
}

Function Check-Housekeeping-Comments {
	Param ([string[]]$housekeeping, [int]$roomNumber);
    $comments = $housekeeping[0].Substring(0, 25);
    if ($comments -ne "                         ") {
        Write-Host "$roomNumber comments: $comments";
    }
}

Function Has-Housekeeping {
	Param ([string[]]$housekeeping);
    $schedule = Parse-Schedule $housekeeping;
    if ($schedule -eq "Overlapping services!") {
        return $true;
    }
	if ((Parse-Days-Count $housekeeping) -ne $schedule.Count) {
		throw "Expected days and schedule to be the same length";
	}
	if (($schedule.Count -lt 9) -and ($schedule[-1] -ne "C/O ")) {
        # Even if it's all spaces,
        # we count deviations from the expected pattern of a trailing "C/O "
		return $true;
	}
	if ($schedule[-1] -eq "C/O ") {
		$schedule = Skip-Last $schedule;
	}
    return Array-Some $schedule {
        Param([string]$x);
        return $x -ne "    ";
    }
}

Function Fill-Tidys {
    Send-Keys-Sequentially "A,T,{F1},~,{F10}";
}

Function Add-First-Rfsh {
    Send-Keys-Sequentially "A,R,{F1},~,N,{F10}";
    Send-Keys-Sequentially "{RIGHT},{RIGHT},{F2},{UP},{UP},{F2},{F10}";
}

Function Add-Housekeeping {
    Param ([int]$scheduleCount);
	if ($scheduleCount -eq 0) {
	    throw "Someone who's about to check out doesn't need housekeeping";
    } elseif ($scheduleCount -le 2) {
        Fill-Tidys;
    } elseif ($scheduleCount -le 6) {
        Fill-Tidys;
        Add-First-Rfsh;
    } elseif ($scheduleCount -eq 7) {
        Fill-Tidys;
        Send-Keys-Sequentially "A,R,{F1},~,N,{F10}";
        Send-Keys-Sequentially "{RIGHT},{RIGHT},{F2}";
        Send-Keys-Sequentially "{UP},{UP},{UP},{F2}";
        Send-Keys-Sequentially "{RIGHT},{RIGHT},{RIGHT},{RIGHT},{F2},{F10}";
    } elseif ($scheduleCount -le 9) {
        Fill-Tidys;
        Send-Keys-Sequentially "A,R,{F1},~,N,{F10}";
        Send-Keys-Sequentially "{RIGHT},{RIGHT},{F2},Y";
        Send-Keys-Sequentially "M,{RIGHT},{RIGHT},{F2},Y";
        Send-Keys-Sequentially "M,{RIGHT},{RIGHT},{RIGHT},{RIGHT},{RIGHT},{RIGHT},{F2},Y";
    } else {
        throw "Unimplemented";
    }
}

Function Add-Housekeeping-If-None {
	$services = Parse-Services $housekeeping;
	if (Are-Services-Weird $services) {
		return "weird services";
	}
    $schedule = Parse-Schedule $housekeeping;
    if ($schedule -eq "Overlapping services!") {
        return "weird overlapping services";
    }
	if ((Parse-Days-Count $housekeeping) -ne $schedule.Count) {
		throw "Expected days and schedule to be the same length";
	}
	if (($schedule.Count -lt 9) -and ($schedule[-1] -ne "C/O ")) {
		return "weird checkout";
	}
	if ($schedule[-1] -eq "C/O ") {
		$schedule = Skip-Last $schedule;
	}
    if ($schedule.Count -eq 0) {
        return "normal";
    }
	if (Are-Non-Checkouts-Weird $schedule) {
		return "weird schedule";
	}
	if (!(Is-Schedule-Empty $schedule)) {
		return "normal";
	}
	if (("TIDY" -in $services) -or ("RFSH" -in $services)) {
	    return "weird unused services";
	}
    Send-Keys "S";
	Add-Housekeeping $schedule.Count;
    Send-Keys "{F4}";
    return "added housekeeping";
}

Function Process-Room {
	Param ([int]$roomNumber);
    $hasJ8 = Has-J8;
	Send-Keys "g";
    $housekeeping = Copy-Housekeeping-Screen;
    Check-Housekeeping-Comments $housekeeping $roomNumber;
    if ($hasJ8) {
        if (Has-Housekeeping $housekeeping) {
            Write-Host "$roomNumber declined housekeeping, but has housekeeping";
        } else {
            Write-Host "$roomNumber declined housekeeping";
        }
    } else {
	    $status = Add-Housekeeping-If-None $housekeeping;
        Write-Host "$roomNumber $status";
    }
	Send-Keys "{F4}";
	Send-Keys "{F4}";
}

if ($foundRooms -eq $null) {
    $foundRooms = [System.Collections.ArrayList]@();
}

Function Skip-Room {
    $null = $Global:foundRooms.Add($false);
}

Function Main {
    Param([int]$startRoom);
    $roomNumbers = $(101..103; 105; 126..129; 201..214; 216..229; 231; 301..329; 331; 401..429; 431);
    if ($Global:foundRooms.Count -lt $roomNumbers.IndexOf($startRoom)) {
        throw "Haven't processed $($roomNumbers[$Global:foundRooms.Count]) yet! Type `"Skip-Room`" to skip it.";
    } elseif ($roomNumbers.IndexOf($startRoom) -eq -1) {
        throw "$startRoom is not a valid room number!";
    }
    Move-Mouse 10 385;
    Left-Click;
    for ($roomIndex = $roomNumbers.IndexOf($startRoom); $roomIndex -lt $roomNumbers.Count; $roomIndex++) {
        $roomNumber = $roomNumbers[$roomIndex];
        $foundRoom = Navigate-To-Room-Number $roomNumber;
        if ($foundRoom) {
            Process-Room $roomNumber;
        } else {
            Write-Host "$roomNumber not found";
        }
        # Don't record a room as done until we process it
        if ($Global:foundRooms.Count -lt $roomIndex) {
            throw "Unreachable branch!";
        } elseif ($Global:foundRooms.Count -eq $roomIndex) {
            $null = $Global:foundRooms.Add($foundRoom);
        } elseif (0 -le $roomIndex) { 
            if ($Global:foundRooms[$roomIndex] -ne $foundRoom) {
                Write-Host "$roomNumber found room on one run but not on another run";
                $Global:foundRooms[$roomIndex] = $foundRoom;
            }
        } else {
            throw "Unreachable branch!";
        }
    }
}
