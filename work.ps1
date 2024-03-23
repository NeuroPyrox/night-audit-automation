$inspect = $null;

# TODO minimize waiting time once everything is implemented
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
        $Global:inspect = $result;
        throw "Please implement one more retry because it worked."
    }
    Wait;
    $result = Get-Clipboard;
    if ($result.Length -ne 0) {
        $Global:inspect = $result;
        throw "Please implement one more retry because it worked."
    }
    throw "Expected a non-empty result";
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
	$found = Retry-Get-Clipboard;
	if ($found -eq "NO MATCHES!                         ") {
		Send-Keys "{F4}";
		return $false;
	}
    try {
        $first3 = $found.Substring(0, 3);
    } catch {
        $Global:inspect = @($found, (Get-Clipboard));
        throw "Inspect this error";
    }
	if ($first3 -eq $roomNumber.ToString()) {
	    Send-Keys "~";
	    return $true;
	}
    if ($first3 -ne "C/O") {
	    throw "Expected to find a checked out room if the room number doesn't match";
    }
	Move-Mouse 710 280;
	Down-Mouse;
	Move-Mouse 1310 280;
	Up-Mouse;
	Right-Click;
	Move-Mouse 1250 290;
	Left-Click;
	$found = Get-Clipboard;
	if ($found.Substring(0, 3) -eq "   ") {
		Send-Keys "{F4}";
		return $false;
	}
	if ($found.Substring(0, 3) -eq $roomNumber.ToString()) {
	    Send-Keys "~";
	    return $true;
	}
    if ($found.Substring(0, 3) -eq "C/O") {
	    throw "Unhandled case of 2 or more checkouts";
    }
	throw "Expected either spaces, the same room number, or `"C/O`"";
}

Function Copy-Housekeeping-Screen {
	Move-Mouse 10 385;
	Down-Mouse;
	Move-Mouse 1240 660;
	Up-Mouse;
	Right-Click;
	Move-Mouse 1250 400;
	Left-Click;
    $clip = Retry-Get-Clipboard;
    $result = $clip -split "`n";
    try {
        $serviceDateCheck = $result[1].Substring(1, 12);
    } catch {
        $Global:inspect = @($clip, $result, (Get-Clipboard));
        throw "Please inspect this error";
    }
	if ($serviceDateCheck -ne "Service Date") {
		throw "Not on the housekeeping screen";
	}
	return Write-Output -NoEnumerate $result;
}

Function Send-Keys-Sequentially {
    Param ([string]$keys);
    ($keys -split ",") | % {Send-Keys $_;};
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
    # TODO implement more cases
    # TODO make robust
	if ($scheduleCount -eq 0) {
	    throw "Someone who's about to check out doesn't need housekeeping";
	} elseif ($scheduleCount -eq 1) {
        Fill-Tidys;
    } elseif ($scheduleCount -eq 2) {
        Fill-Tidys;
    } elseif ($scheduleCount -eq 3) {
        Fill-Tidys;
        Add-First-Rfsh;
    } elseif ($scheduleCount -eq 4) {
        Fill-Tidys;
        Add-First-Rfsh;
    } elseif ($scheduleCount -eq 5) {
        Fill-Tidys;
        Add-First-Rfsh;
    } elseif ($scheduleCount -eq 7) {
        Fill-Tidys;
        Send-Keys-Sequentially "A,R,{F1},~,N,{F10}";
        Send-Keys-Sequentially "{RIGHT},{RIGHT},{F2}";
        Send-Keys-Sequentially "{UP},{UP},{UP},{F2}";
        Send-Keys-Sequentially "{RIGHT},{RIGHT},{RIGHT},{RIGHT},{F2},{F10}";
    } elseif ($scheduleCount -eq 9) {
        Fill-Tidys;
        Send-Keys-Sequentially "A,R,{F1},~,N,{F10}";
        Send-Keys-Sequentially "{RIGHT},{RIGHT},{F2},Y";
        Send-Keys-Sequentially "M,{RIGHT},{RIGHT},{F2},Y";
        Send-Keys-Sequentially "M,{RIGHT},{RIGHT},{RIGHT},{RIGHT},{RIGHT},{RIGHT},{F2},Y";
    } else {
        throw "Unimplemented";
    }
}

Function Skip-Last {
    Param ([object[]]$array);
    if ($array.Count -eq 0) {
        throw "Can't remove from an empty array"
    }
    if ($array.Count -eq 1) {
        return Write-Output -NoEnumerate @();
    }
	return Write-Output -NoEnumerate $array[-$schedule.Count..-2];
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

Function Parse-Services {
	Param ([string[]]$housekeeping);
	$services = 5..9 | % {
		if ($housekeeping[$_].Length -ne 72) {
			throw "Expected 72 characters";
		}
		$housekeeping[$_].Substring(1, 17);
	}
	$services = Trim-End $services {Param($x); $x -eq "                 ";};
	if ($services.Count -ne ($services | Select-Object -Unique).Count) {
		throw "Should only have one of each type of available service";
	}
	return Write-Output -NoEnumerate $services;
}

Function Parse-Days-Count {
	Param ([string[]]$housekeeping);
	$days = $housekeeping[0];
	if ($days.Length -ne 72) {
		throw "Expected 72 characters";
	}
	$days = 0..8 | % {$days.Substring(21 + (6 * $_), 3)};
	$days = Trim-End $days {Param($x); $x -eq "   "};
	if (Array-Some $days {Param([string]$x); !($x -in "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");}) {
		throw "Unexpected value for a day";
	}
	return $days.Count;
}

Function Print-Schedule {
	Param ([string[][]]$schedule);
    $schedule | % {
        Write-Host $_.Count, $_;
    }
}

# Assumes services aren't weird
Function Parse-Schedule {
	Param ([string[]]$housekeeping);
    $schedule = [System.Collections.ArrayList]@();
	0..8 | % {
		$dayIndex = $_;
        $dayServices = [System.Collections.ArrayList]@();
		5..9 |
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
    $trimmed = Trim-End $schedule {Param([string[]]$x); $x.Count -eq 0};
    if ("C/O " -in $trimmed[-1]) {
	    return Write-Output -NoEnumerate $trimmed;
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

Function Is-Checkout-Weird {
	Param ([string[][]]$schedule);
	if (Array-Some (Skip-Last $schedule) {Param($x); "C/O " -in $x}) {
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
                if ((1 -lt $schedule[$dayIndex].Count) -or !( `
                    (($weeklyPattern[$dayIndex] -eq 0) -and ($schedule[$dayIndex].Count -eq 0)) `
                    -or (($weeklyPattern[$dayIndex] -eq 1) -and ("TIDY" -in $schedule[$dayIndex])) `
                    -or (($weeklyPattern[$dayIndex] -eq 2) -and ("RFSH" -in $schedule[$dayIndex])) `
                    -or (($weeklyPattern[$dayIndex] -eq 3) -and ("1XWE" -in $schedule[$dayIndex])) `
                )) {
                    return $false;
                }
			}
			return $true;
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
    $Global:inspect = $housekeeping;
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
		return Write-Host "$roomNumber weird C/O";
	}
	if ("C/O " -in $schedule[-1]) {
		$schedule = Skip-Last $schedule;
	}
    if ($schedule.Count -eq 0) {
        return Write-Host "$roomNumber normal"
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
    Send-Keys "S";
	Add-Housekeeping $schedule.Count
    Send-Keys "{F4}";
    # TODO double check Add-Housekeeping
    Write-Host "$roomNumber added housekeeping"
}

# TODO retry on errors
Function Main {
    Param([int]$startRoom)
    $roomNumbers = $(101..103; 105; 126..129; 201..214; 216..229; 231; 301..329; 331; 401..429; 431);
    Move-Mouse 10 385;
    Left-Click;
    for ($roomIndex = $roomNumbers.IndexOf($startRoom); $roomIndex -lt $roomNumbers.Count; $roomIndex++) {
        $roomNumber = $roomNumbers[$roomIndex];
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
