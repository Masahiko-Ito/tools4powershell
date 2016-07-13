#
# pscat - concatenate files and print on the standard output
#
function pscat {
	begin{
		$numberSw = "off"
		$number = 0
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				write-output "Usage: pscat [-h|--help] [-n] [input ...]"
				write-output "Concatenate input(s), or standard input, to standard output."
				write-output ""
				write-output "  -n        number all output lines"
				return
			}elseif ($args[$i] -eq "-n"){
				$numberSw = "on"
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				if ($numberSw -eq "off"){
					get-content $files[$i]
				}else{
					get-content $files[$i] |
					foreach-object {
						$number++
						$out = $number.tostring() + " " + $_
						write-output $out
					}
				}
			}
		}else{
			if ($numberSw -eq "off"){
				write-output $_
			}else{
				$number++
				$out = $number.tostring() + " " + $_
				write-output $out
			}
		}
	}
	end{
	}
}

#
# psgrep - print lines matching a pattern
#
function psgrep {
	begin{
		$ignorecaseSw = "off"
		$invertSw = "off"
		$string = ""
		$stringSw = "off"
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				write-output "Usage: psgrep [-h|--help] [-v] [-i] regex [input ...]"
				write-output "Search for regex in each input or standard input."
				write-output ""
				write-output "  -v        select non-matching lines"
				write-output "  -i        ignore case distinctions"
				return
			}elseif ($args[$i] -eq "-v"){
				$invertSw = "on"
			}elseif ($args[$i] -eq "-i"){
				$ignorecaseSw = "on"
			}else{
				if ($stringSw -eq "off"){
					$string = $args[$i]
					$stringSw = "on"
				}else{
					$files[$filesIndex] = $args[$i]
					$filesIndex++
				}
			}
		}
	}
	process{
		if ($invertSw -eq "off"){
			if ($filesIndex -gt 0){
				for ($i = 0; $i -lt $filesIndex; $i++){
					get-content $files[$i] |
					foreach-object {
						if ($ignorecaseSw -eq "off"){
							if ($_ -cmatch $string){
								$out = $files[$i] + ":" + $_
								write-output $out
							}
						}else{
							if ($_ -match $string){
								$out = $files[$i] + ":" + $_
								write-output $out
							}
						}
					}
				}
			}else{
				if ($ignorecaseSw -eq "off"){
					if ($_ -cmatch $string){
						write-output $_
					}
				}else{
					if ($_ -match $string){
						write-output $_
					}
				}
			}
		}else{
			if ($filesIndex -gt 0){
				for ($i = 0; $i -lt $filesIndex; $i++){
					get-content $files[$i] |
					foreach-object {
						if ($ignorecaseSw -eq "off"){
							if ($_ -cnotmatch $string){
								$out = $files[$i] + ":" + $_
								write-output $out
							}
						}else{
							if ($_ -notmatch $string){
								$out = $files[$i] + ":" + $_
								write-output $out
							}
						}
					}
				}
			}else{
				if ($ignorecaseSw -eq "off"){
					if ($_ -cnotmatch $string){
						write-output $_
					}
				}else{
					if ($_ -notmatch $string){
						write-output $_
					}
				}
			}
		}
	}
	end{
	}
}

#
# pswcl - print newline counts for each file
#
function pswcl {
	begin{
		$helpSw = $false
		$number = 0
		$total_number = 0
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: pswcl [-h|--help] [input ...]"
				write-output "Print newline counts for each input, and a total line if more than one input is specified."
				return
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				$number = 0
				get-content $files[$i] |
				foreach-object {
					$number++
				}
				$out = $number.tostring() + " " + $args[$i]
				write-output $out
				$total_number += $number
			}
		}else{
			if ($_ -ne $null){
				$number++
			}
		}
	}
	end{
		if ($helpSw -eq $false){
			if ($filesIndex -eq 0){
				write-output $number
			}else{
				$out = $total_number.tostring() + " TOTAL"
				write-output $out
			}
		}
	}
}

#
# pssed - stream editor for filtering and transforming text
#
function pssed {
	begin{
		$before_string = $null
		$after_string = $null
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				write-output "Usage: pssed [-h|--help] regex string [input ...]"
				write-output "For each substring matching regex in each lines from input, substitute the string."
				return
			}elseif ($before_string -eq $null){
				$before_string = $args[$i]
			}elseif ($after_string -eq $null){
				$after_string = $args[$i]
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				get-content $files[$i] |
				foreach-object {
					$out = $_ -replace $before_string, $after_string
					write-output $out
				}
			}
		}else{
			$out = $_ -replace $before_string, $after_string
			write-output $out
		}
	}
	end{
	}
}

#
# pshead - output the first part of files
#
function pshead {
	begin{
		$line = 10
		$wline = 0
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				write-output "Usage: pshead [-h|--help] [-l line_number] [input ...]"
				write-output "Print the first 10 lines of each input to standard output."
				write-output "With no input, read standard input."
				write-output ""
				write-output "  -l line_number        print the first line_number lines instead of the first 10"
				return
			}elseif ($args[$i] -eq "-l"){
				$i++
				$line = $args[$i]
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				$wline = 0
				get-content $files[$i] |
				select-object -first $line
#				foreach-object {
#					if ($wline -lt $line){
#						write-output $_
#					}else{
#						break
#					}
#					$wline++
#				}
			}
		}else{
			if ($wline -lt $line){
				write-output $_
			}else{
				break
			}
			$wline++
		}
	}
	end{
	}
}

#
# pstail - output the last part of files
#
function pstail {
	begin{
		$helpSw = $false
		$tmpfile = [System.IO.Path]::GetTempFileName()
		$line = 10
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: pstail [-h|--help] [-l line_number] [input ...]"
				write-output "Print the last 10 lines of each input to standard output."
				write-output "With no input, read standard input."
				write-output ""
				write-output "  -l line_number        print the last line_number lines instead of the last 10"
				return
			}elseif ($args[$i] -eq "-l"){
				$i++
				$line = $args[$i]
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				get-content $files[$i] |
				select-object -last $line
			}
		}else{
			write-output $_ >>$tmpfile
		}
	}
	end{
		if ($helpSw -eq $false){
			if ($filesIndex -eq 0){
				get-content $tmpfile |
				select-object -last $line
				remove-item $tmpfile
			}
		}
	}
}

#
# pscut - remove sections from each line of files
#
function pscut {
	begin{
		$delimiter = ","
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				write-output "Usage: pscut [-h|--help] [-d ""delimiter""] -i ""index,..."" [input ...]"
				write-output "Print selected parts of lines from each input to standard output."
				write-output "With no input read standard input."
				write-output ""
				write-output "  -d ""delimiter""        use ""delimiter"" instead of "","" for field delimiter"
				write-output "  -i ""index,...""        select only these fields(0 origin)"
				return
			}elseif ($args[$i] -eq "-d"){
				$i++
				$delimiter = $args[$i]
			}elseif ($args[$i] -eq "-i"){
				$i++
				$cols = $args[$i] -split ","
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				get-content $files[$i] |
				foreach-object {
					$out = ""
					foreach ($j in $cols){
						if ($out -eq ""){
							$out = $_.split($delimiter)[$j]
						}else{
							$out = $out + $delimiter + $_.split($delimiter)[$j]
						}
					}
					write-output $out
				}
			}
		}else{
			$out = ""
			if ($_ -ne $null){
				foreach ($j in $cols){
					if ($out -eq ""){
						$out = $_.split($delimiter)[$j]
					}else{
						$out = $out + $delimiter + $_.split($delimiter)[$j]
					}
				}
				write-output $out
			}
		}
	}
	end{
	}
}

#
# pstee - read from standard input and write to standard output and files
#
function pstee {
	begin{
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				write-output "Usage: pstee [-h|--help] [output ...]"
				write-output "Copy standard input to each output, and also to standard output."
				return
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		write-output $_
		for ($i = 0; $i -lt $filesIndex; $i++){
			write-output $_ >>$files[$i]
		}
	}
	end{
	}
}

#
# psuniq - report or omit repeated lines
#
function psuniq {
	begin{
		$helpSw = $false
		$oldrec = $null
		$count = 0
		$duplicateSw = "off"
		$countSw = "off"
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: psuniq [-h|--help] [-d|-c] [input ...]"
				write-output "Filter adjacent matching lines from input (or standard input),"
				write-output "writing to standard output."
				write-output "With no options, matching lines are merged to the first occurrence."
				write-output ""
				write-output "  -d        only print duplicate lines"
				write-output "  -c        prefix lines by the number of occurrences"
				return
			}elseif ($args[$i] -eq "-d"){
				$duplicateSw = "on"
			}elseif ($args[$i] -eq "-c"){
				$countSw = "on"
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($filesIndex -gt 0){
			for ($i = 0; $i -lt $filesIndex; $i++){
				$oldrec = $null
				$count = 0
				get-content $files[$i] |
				foreach-object {
					if ($duplicateSw -eq "on"){
						if ($oldrec -eq $null){
							$oldrec = $_
							$isDuplicateSw = "off"
						}else{
							if ($_ -eq $oldrec){
								$isDuplicateSw = "on"
							}else{
								if ($isDuplicateSw -eq "on"){
									write-output $oldrec
								}
								$oldrec = $_	
								$isDuplicateSw = "off"
							}
						}
					}elseif ($countSw -eq "on"){
						if ($oldrec -eq $null){
							$oldrec = $_
							$count++
						}else{
							if ($_ -eq $oldrec){
								$count++
							}else{
								$out = $count.tostring() + " " + $oldrec
								write-output $out
								$count = 0
								$oldrec = $_	
								$count++
							}
						}
					}else{
						if ($oldrec -eq $null){
							$oldrec = $_
							$isDuplicateSw = "off"
						}else{
							if ($_ -eq $oldrec){
								$isDuplicateSw = "on"
							}else{
								if ($isDuplicateSw -eq "off"){
									write-output $oldrec
								}
								$oldrec = $_	
								$isDuplicateSw = "off"
							}
						}
					}
				}
			}
		}else{
			if ($duplicateSw -eq "on"){
				if ($oldrec -eq $null){
					$oldrec = $_
					$isDuplicateSw = "off"
				}else{
					if ($_ -eq $oldrec){
						$isDuplicateSw = "on"
					}else{
						if ($isDuplicateSw -eq "on"){
							write-output $oldrec
						}
						$oldrec = $_	
						$isDuplicateSw = "off"
					}
				}
			}elseif ($countSw -eq "on"){
				if ($oldrec -eq $null){
					$oldrec = $_
					$count++
				}else{
					if ($_ -eq $oldrec){
						$count++
					}else{
						$out = $count.tostring() + " " + $oldrec
						write-output $out
						$count = 0
						$oldrec = $_	
						$count++
					}
				}
			}else{
				if ($oldrec -eq $null){
					$oldrec = $_
					$isDuplicateSw = "off"
				}else{
					if ($_ -eq $oldrec){
						$isDuplicateSw = "on"
					}else{
						if ($isDuplicateSw -eq "off"){
							write-output $oldrec
						}
						$oldrec = $_	
						$isDuplicateSw = "off"
					}
				}
			}
		}
	}
	end{
		if ($helpSw -eq $false){
			if ($duplicateSw -eq "on"){
				if ($isDuplicateSw -eq "on"){
					write-output $oldrec
				}
			}elseif ($countSw -eq "on"){
				$out = $count.tostring() + " " + $oldrec
				write-output $out
			}else{
				if ($isDuplicateSw -eq "off"){
					write-output $oldrec
				}
			}
		}
	}
}

#
# psjoin - join lines of two files on a common field
#
function psjoin {
	begin{
		$helpSw = $false
		$delimiter = ","
		$keyidx1 = "0"
		$keyidx2 = "0"
		$action = "m"
		$multikey = ""
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: psjoin [-h|--help] [-d ""delimiter""] [-1 ""index,...""] [-2 ""index,...""] [-a [m|1|2|12|21]] [-m [1|2]] input1 input2"
				write-output "For each pair of input lines with identical join fields, write a line to"
				write-output "standard output.  The default join field is the first, delimited by "",""."
				write-output ""
				write-output "  -d ""delimiter""        use ""delimiter"" as input and output field separator instead of "","""
				write-output "  -1 ""index,...""        join on this index(s) of file 1 (default 0)"
				write-output "  -2 ""index,...""        join on this index(s) of file 2 (default 0)"
				write-output "  -a m                    write only matching lines (default)"
				write-output "     1                    write only unpairable lines from input1"
				write-output "     2                    write only unpairable lines from input2"
				write-output "     12                   write all lines from input1 and matching lines from input2"
				write-output "     21                   write all lines from input2 and matching lines from input1"
				write-output "  -m [1|2]                specify input which has multiple join fields"
				return
			}elseif ($args[$i] -eq "-d"){
				$i++
				$delimiter = $args[$i]
			}elseif ($args[$i] -eq "-1"){
				$i++
				$keyidx1 = $args[$i]
			}elseif ($args[$i] -eq "-2"){
				$i++
				$keyidx2 = $args[$i]
			}elseif ($args[$i] -eq "-a"){
				$i++
				$action = $args[$i]
			}elseif ($args[$i] -eq "-m"){
				$i++
				$multikey = $args[$i]
			}else{
				$files[$filesIndex] = resolve-path $args[$i]
				$filesIndex++
			}
		}
#		$oIn1 = New-Object System.IO.StreamReader($files[0],[Text.Encoding]::GetEncoding("Shift_JIS"))
		$oIn1 = New-Object System.IO.StreamReader($files[0],[Text.Encoding]::Default)
		$oIn2 = New-Object System.IO.StreamReader($files[1],[Text.Encoding]::Default)
	}
	process{
		if ($helpSw -eq $false){
			function sub_start {
				$global:endSw = "off"
				$global:rec1 = $oIn1.readLine()
				$global:rec2 = $oIn2.readLine()
				$global:matchkey = $null
			}
			function sub_main {
				$key1 = mkkey $global:rec1 $keyidx1
				$key2 = mkkey $global:rec2 $keyidx2
				if ($action -eq "1"){
					if (islt $key1 $key2){
						if ($global:matchkey -eq $null -or $global:matchkey -cne $key1){
							write-output $global:rec1
						}
						$global:rec1 = $oIn1.readLine()
					}elseif (islt $key2 $key1){
						$global:rec2 = $oIn2.readLine()
					}else{
						$global:matchkey = $key1
						if ($multikey -eq 1){
							$global:rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$global:rec2 = $oIn2.readLine()
						}else{
							$global:rec1 = $oIn1.readLine()
							$global:rec2 = $oIn2.readLine()
						}
					}
				}elseif ($action -eq "2"){
					if (islt $key2 $key1){
						if ($global:matchkey -eq $null -or $global:matchkey -cne $key2){
							write-output $global:rec2
						}
						$global:rec2 = $oIn2.readLine()
					}elseif (islt $key1 $key2){
						$global:rec1 = $oIn1.readLine()
					}else{
						$global:matchkey = $key2
						if ($multikey -eq 1){
							$global:rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$global:rec2 = $oIn2.readLine()
						}else{
							$global:rec1 = $oIn1.readLine()
							$global:rec2 = $oIn2.readLine()
						}
					}
				}elseif ($action -eq "12"){
					if (islt $key1 $key2){
						if ($global:matchkey -eq $null -or $global:matchkey -cne $key1){
							write-output $global:rec1
						}
						$global:rec1 = $oIn1.readLine()
					}elseif (islt $key2 $key1){
						$global:rec2 = $oIn2.readLine()
					}else{
						$out = $global:rec1 + $delimiter + $global:rec2
						write-output $out
						$global:matchkey = $key1
						if ($multikey -eq 1){
							$global:rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$global:rec2 = $oIn2.readLine()
						}else{
							$global:rec1 = $oIn1.readLine()
							$global:rec2 = $oIn2.readLine()
						}
					}
				}elseif ($action -eq "21"){
					if (islt $key2 $key1){
						if ($global:matchkey -eq $null -or $global:matchkey -cne $key2){
							write-output $global:rec2
						}
						$global:rec2 = $oIn2.readLine()
					}elseif (islt $key1 $key2){
						$global:rec1 = $oIn1.readLine()
					}else{
						$out = $global:rec2 + $delimiter + $global:rec1
						write-output $out
						$global:matchkey = $key2
						if ($multikey -eq 1){
							$global:rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$global:rec2 = $oIn2.readLine()
						}else{
							$global:rec1 = $oIn1.readLine()
							$global:rec2 = $oIn2.readLine()
						}
					}
				}else{
					if (islt $key1 $key2){
						$global:rec1 = $oIn1.readLine()
					}elseif (islt $key2 $key1){
						$global:rec2 = $oIn2.readLine()
					}else{
						$out = $global:rec1 + $delimiter + $global:rec2
						write-output $out
						if ($multikey -eq 1){
							$global:rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$global:rec2 = $oIn2.readLine()
						}else{
							$global:rec1 = $oIn1.readLine()
							$global:rec2 = $oIn2.readLine()
						}
					}
				}
				if ($global:rec1 -eq $null -and $global:rec2 -eq $null){
					$global:endSw = "on"
				}
			}
			function sub_end {
			}
			function mkkey($rec, $key){
				if ($rec -eq $null){
					$ret = $null
				}else{
					$ret = ""
					$recs = $rec -split $delimiter
					$keys = $key -split ","
					foreach ($i in $keys){
						$ret = $ret.tostring() + $recs[$i]
					}
				}
				return $ret
			}
			function islt($key1, $key2){
				if ($key1 -eq $null -and $key2 -eq $null){
					$ret = $false
				}elseif ($key1 -eq $null){
					$ret = $false
				}elseif ($key2 -eq $null){
					$ret = $true
				}else{
					if ($key1.tostring() -clt $key2.tostring()){
						$ret = $true
					}else{
						$ret = $false
					}
				}
				return $ret
			}
	
			sub_start
			while ($global:endSw -eq "off"){
				sub_main
			}
			sub_end
		}
	}
	end{
		if ($helpSw -eq $false){
			$oIn1.close()
			$oIn2.close()
		}
	}
}

#
# psxls2csv - convert excel to csv
#
function psxls2csv {
	begin{
		$helpSw = $false
		$strInput = ""
		$sheet = 1
		$strOutput = ""
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: psxls2csv [-h|--help] [-i input] [-s sheet] [-o [output|-]]"
				write-output "Convert specified excel sheet to csv."
				write-output "If input is not specified, all excel files in current directory will be converted."
				write-output "If output is not specified, input will be converted into same filename, but with extention "".csv""."
				write-output "If ""-"" is specified for ""-o"" option, input will be converted into stdout."
				write-output ""
				write-output "BUGS"
				write-output "  If input is not specified and output is specified, only last excel sheet in current directory will remain in output file."
				return
			}elseif ($args[$i] -eq "-i"){
				$i++
				$strInput = $args[$i]
			}elseif ($args[$i] -eq "-s"){
				$i++
				$sheet = $args[$i]
			}elseif ($args[$i] -eq "-o"){
				$i++
				$strOutput = $args[$i]
			}else{
				$files[$filesIndex] = $args[$i]
				$filesIndex++
			}
		}
	}
	process{
		if ($helpSw -eq $true){
			return
		}
		if ($strInput -eq ""){
			Get-ChildItem *.xls* |
			ForEach-Object {
				$InPath = resolve-path $_.Name
				if ($strOutput -eq ""){
					$OutPath = $InPath -replace ".xls.*", ".csv"
					$out = $_.Name + " -> " + ($_.Name -replace ".xls.*", ".csv")
					write-output $out
				}elseif ($strOutput -eq "-"){
					$OutPath = [System.IO.Path]::GetTempFileName()
				}else{
					$OutPath = (get-location).tostring() + "\" + $strOutput
				}
				$objExcel = New-Object -ComObject Excel.Application
				$objExcel.DisplayAlerts = $false
			
				$objExcel.Workbooks.open($InPath) | out-null
			
				$objSheet = $objExcel.Worksheets.Item($sheet)
				$objSheet.SaveAs($OutPath, 6)
				$objExcel.Workbooks.Close()
				$objExcel.Quit()
				if ($strOutput -eq "-"){
					get-content $OutPath
					remove-item $OutPath
				}
			}
		}else{
			$InPath = resolve-path $strInput
			if ($strOutput -eq ""){
				$OutPath = $InPath -replace ".xls.*", ".csv"
			}elseif ($strOutput -eq "-"){
				$OutPath = [System.IO.Path]::GetTempFileName()
			}else{
				$OutPath = (get-location).tostring() + "\" + $strOutput
			}
			$objExcel = New-Object -ComObject Excel.Application
			$objExcel.DisplayAlerts = $false
		
			$objExcel.Workbooks.open($InPath) | out-null
		
			$objSheet = $objExcel.Worksheets.Item($sheet)
			$objSheet.SaveAs($OutPath, 6)
			$objExcel.Workbooks.Close()
			$objExcel.Quit()
			if ($strOutput -eq "-"){
				get-content $OutPath
				remove-item $OutPath
			}
		}
	}
	end{
	}
}
