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
								if ($filesIndex -eq 1){
									$out = $_
								}else{
									$out = $files[$i] + ":" + $_
								}
								write-output $out
							}
						}else{
							if ($_ -match $string){
								if ($filesIndex -eq 1){
									$out = $_
								}else{
									$out = $files[$i] + ":" + $_
								}
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
								if ($filesIndex -eq 1){
									$out = $_
								}else{
									$out = $files[$i] + ":" + $_
								}
								write-output $out
							}
						}else{
							if ($_ -notmatch $string){
								if ($filesIndex -eq 1){
									$out = $_
								}else{
									$out = $files[$i] + ":" + $_
								}
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
				if ($filesIndex -gt 1){
					$out = $total_number.tostring() + " TOTAL"
					write-output $out
				}
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
				write-output "With no input, read standard input."
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
#								if ($isDuplicateSw -eq "off"){
									write-output $oldrec
#								}
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
#						if ($isDuplicateSw -eq "off"){
							write-output $oldrec
#						}
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
#				if ($isDuplicateSw -eq "off"){
					write-output $oldrec
#				}
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
#				$files[$filesIndex] = (resolve-path $args[$i]).Path
				$files[$filesIndex] = psabspath $args[$i]
				$InPath = psabspath $_.Name
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
		$tabSw = "off"
		$strOutput = ""
		$files = @{}
		$filesIndex = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: psxls2csv [-h|--help] [-i input] [-s sheet] [-t] [-o [output|-]]"
				write-output "Convert specified excel sheet to csv or tsv with -t option."
				write-output "If input is not specified, all excel files in current directory will be converted."
				write-output "If output is not specified, input will be converted into same filename, but with extention "".csv"" or "".txt""."
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
			}elseif ($args[$i] -eq "-t"){
				$tabSw = "on"
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
#				$InPath = (resolve-path $_.Name).Path
				$InPath = psabspath $_.Name
				if ($strOutput -eq ""){
					if ($tabSw -eq "on"){
						$OutPath = $InPath -replace ".xls.*", ".txt"
						$out = $_.Name + " -> " + ($_.Name -replace ".xls.*", ".txt")
					}else{
						$OutPath = $InPath -replace ".xls.*", ".csv"
						$out = $_.Name + " -> " + ($_.Name -replace ".xls.*", ".csv")
					}
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
				if ($tabSw -eq "on"){
#					$objSheet.SaveAs($OutPath, -4158)
					$objSheet.SaveAs($OutPath, 42)
				}else{
					$objSheet.SaveAs($OutPath, 6)
				}
				$objExcel.Workbooks.Close()
				$objExcel.Quit()
				if ($tabSw -eq "on"){
					$TmpOutPath = [System.IO.Path]::GetTempFileName()
					Get-Content $OutPath | Out-File -Encoding default $TmpOutPath 
					Move-Item -Force -path $TmpOutPath -destination $OutPath
				}
				if ($strOutput -eq "-"){
					get-content $OutPath
					remove-item $OutPath
				}
			}
		}else{
			if ($strInput -match '^\\'){
				$InPath = $strInput
			}else{
#				$InPath = (resolve-path $strInput).Path
				$InPath = psabspath $strInput
			}
			if ($strOutput -eq ""){
				if ($tabSw -eq "on"){
					$OutPath = $InPath -replace ".xls.*", ".txt"
					$out = $_.Name + " -> " + ($_.Name -replace ".xls.*", ".txt")
				}else{
					$OutPath = $InPath -replace ".xls.*", ".csv"
					$out = $_.Name + " -> " + ($_.Name -replace ".xls.*", ".csv")
				}
			}elseif ($strOutput -eq "-"){
				$OutPath = [System.IO.Path]::GetTempFileName()
			}else{
				$OutPath = (get-location).tostring() + "\" + $strOutput
			}
			$objExcel = New-Object -ComObject Excel.Application
			$objExcel.DisplayAlerts = $false
		
			$objExcel.Workbooks.open($InPath) | out-null
		
			$objSheet = $objExcel.Worksheets.Item($sheet)
			if ($tabSw -eq "on"){
#				$objSheet.SaveAs($OutPath, -4158)
				$objSheet.SaveAs($OutPath, 42)
			}else{
				$objSheet.SaveAs($OutPath, 6)
			}
			$objExcel.Workbooks.Close()
			$objExcel.Quit()
			if ($tabSw -eq "on"){
				$TmpOutPath = [System.IO.Path]::GetTempFileName()
				Get-Content $OutPath | Out-File -Encoding default $TmpOutPath 
				Move-Item -Force -path $TmpOutPath -destination $OutPath
			}
			if ($strOutput -eq "-"){
				get-content $OutPath
				remove-item $OutPath
			}
		}
	}
	end{
		if ($helpSw -ne $true){
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objSheet) | out-null
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | out-null
			[System.GC]::Collect() | out-null
		}
	}
}

#
# print
#
function psprint (){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psprint [-h|--help] [arg ...]"
			write-output "Print arguments to standard output."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			$out = $args -join ""
			write-output $out
		}
	}
}

#
# pstmpfile - get temporary file
#
function pstmpfile (){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: pstmpfile [-h|--help]"
			write-output "Print temporary file path to standard output."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			[System.IO.Path]::GetTempFileName()
		}
	}
}

#
# psabspath - get absolute file path
#
function psabspath ($path){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psabspath [-h|--help] file"
			write-output "Print absolute file path for file to standard output."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
#			(get-location).tostring() + "\" + $path
			(get-location).ProviderPath + "\" + $path
		}
	}
}

#
# psreadpassword - get password safely from console
#
function psreadpassword (){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psreadpassword [-h|--help]"
			write-output "Get password safely from console and decrypt and print it to console."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			$ss = Read-Host -AsSecureString
			$ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
			$password = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
			[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr);
			Write-Output $password
		}
	}
}

#
# psconwrite - output arguments to console without newline
#
function psconwrite(){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psconwrite [-h|--help] [arg ...]"
			write-output "Print arguments to console without newline."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			$out = $args -join ""
			[Console]::Out.write($out)
		}
	}
	
}

#
# psconwriteline - output arguments to console with newline
#
function psconwriteline(){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psconwriteline [-h|--help] [arg ...]"
			write-output "Print arguments to console with newline."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			$out = $args -join ""
			[Console]::Out.writeLine($out)
		}
	}
	
}

#
# psconreadline - read from console
#
function psconreadline(){
	begin{
		$helpSw = $false
		if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psconreadline [-h|--help]"
			write-output "Read line and print it to console with newline."
			write-output ""
			return
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			$input = [Console]::In.readLine()
			write-output $input
		}
	}
}

#
# psopen - Open IO.Stream
#
# ex. $inObj = psopen -r "input.txt"
#     $outObj = psopen -w "output.txt"
#     while (($rec = $inObj.readLine()) -ne $null){
#         $outObj.writeLine($rec)
#     }
#     $inObj.close()
#     $outObj.close()
#
function psopen(){
	begin{
		$encoding = "Shift_JIS"
		$helpSw = $false
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: psopen [-h|--help] [[-r|-w|-a] [inputfile|output]] [-e encoding]"
				write-output "Open IO.Stream and get object."
				write-output "ex."
				write-output '  $inObj = psopen -r "input.txt"'
				write-output '  $outObj = psopen -w "output.txt"'
				write-output '  while (($rec = $inObj.readLine()) -ne $null){'
				write-output '      $outObj.writeLine($rec)'
				write-output '  }'
				write-output '  $inObj.close()'
				write-output '  $outObj.close()'
				write-output ""
				return
			}elseif ($args[$i] -eq "-r"){
				$iomode = "r"
				$i++
				if ($args[$i] -match "^[A-Za-z]:" -or $args[$i] -match "^\\"){
					$inputfile = $args[$i]
				}else{
					$inputfile = (get-location).tostring() + "\" + $args[$i]
				}
			}elseif ($args[$i] -eq "-w"){
				$iomode = "w"
				$i++
				if ($args[$i] -match "^[A-Za-z]:" -or $args[$i] -match "^\\"){
					$outputfile = $args[$i]
				}else{
					$outputfile = (get-location).tostring() + "\" + $args[$i]
				}
			}elseif ($args[$i] -eq "-a"){
				$iomode = "a"
				$i++
				if ($args[$i] -match "^[A-Za-z]:" -or $args[$i] -match "^\\"){
					$outputfile = $args[$i]
				}else{
					$outputfile = (get-location).tostring() + "\" + $args[$i]
				}
			}elseif ($args[$i] -eq "-e"){
				$i++
				$encoding = $args[$i]
			}
		}
	}
	process{
	}
	end{
		if ($helpSw -eq $false){
			if ($iomode -eq "r"){
				$objIO = New-Object System.IO.StreamReader($inputfile, [Text.Encoding]::GetEncoding($encoding))
			}elseif ($iomode -eq "w"){
				$objIO = New-Object System.IO.StreamWriter($outputfile, $false, [Text.Encoding]::GetEncoding($encoding))
			}else{
				$objIO = New-Object System.IO.StreamWriter($outputfile, $true, [Text.Encoding]::GetEncoding($encoding))
			}
			return $objIO
		}
	}
}

#
# psexcel_open - Open excel_file and get excel_object for it
#
function psexcel_open($xlsPath) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_open excel_file"
		write-output "Open excel_file and get excel_object for it."
		write-output "ex."
		write-output '    $xls = psexcel_open "Foo.xlsx"'
		write-output ""
		return
	}

	$objExcel = New-Object -ComObject Excel.Application

#	$objExcel.Visible = $true
	$objExcel.Visible = $false

#	$objExcel.DisplayAlerts = $true
	$objExcel.DisplayAlerts = $false

	if ($xlsPath -match "^[A-Za-z]:" -or $xlsPath -match "^\\"){
		$strPath = $xlsPath
	}else{
		$strPath = (get-location).tostring() + "\" + $xlsPath
	}

	try {
		$objExcel.Workbooks.Open($strPath) | out-null
	}catch{
		$objExcel.Workbooks.Add() | out-null
		$objExcel.Workbooks.item(1).SaveAs($strPath) | out-null
		$objExcel.Workbooks.Close() | out-null
		$objExcel.Workbooks.Open($strPath) | out-null
	}

	return $objExcel
}

#
# psexcel_update - Save by overwrite
#
function psexcel_update($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_update excel_object"
		write-output "Save by overwrite."
		write-output "ex."
		write-output '    psexcel_update $xls'
		write-output ""
		return
	}

	$objExcel.Save() | out-null
}

#
# psexcel_save - Save into excel_file
#
function psexcel_save($objExcel, $xlsPath) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_save excel_object excel_file"
		write-output "Save into excel_file."
		write-output "ex."
		write-output '    psexcel_save $xls "Foo.xlsx"'
		write-output ""
		return
	}

	if ($xlsPath -match "^[A-Za-z]:" -or $xlsPath -match "^\\"){
		$strPath = $xlsPath
	}else{
		$strPath = (get-location).tostring() + "\" + $xlsPath
	}
	$objExcel.Workbooks.item(1).SaveAs($strPath) | out-null
}

#
# psexcel_close - Close excel file by excel_object
#
function psexcel_close($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_close excel_object"
		write-output "Close excel file by excel_object."
		write-output "ex."
		write-output '    psexcel_close $xls'
		write-output ""
		return
	}

	$objExcel.Workbooks.Close() | out-null
	$objExcel.Quit() | out-null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | out-null
	[System.GC]::Collect() | out-null
}

#
# psexcel_getCell - Get value from range on sheet
#
function psexcel_getCell($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getCell excel_object sheet range"
		write-output "Get value from range on sheet."
		write-output "ex."
		write-output '    $val1 = psexcel_getCell $xls "Sheet1" "A1"'
		write-output '    $val2 = psexcel_getCell $xls 2 "B2"'
		write-output ""
		return
	}

	return $objExcel.Worksheets.Item($sheet).Range($range).Text
}

#
# psexcel_setCell Set value on range on sheet
#
function psexcel_setCell($objExcel, $sheet, $range, $text) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_setCell excel_object sheet range value"
		write-output "Set value on range on sheet."
		write-output "ex."
		write-output '    psexcel_setCell $xls "Sheet1" "A1" "some text"'
		write-output '    psexcel_setCell $xls 2 "C3" "=A1+B2"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).Value2 = $text
}

#
# psexcel_getFormula - Get formula from range on sheet
#
function psexcel_getFormula($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getFormula excel_object sheet range"
		write-output "Get formula from range on sheet."
		write-output "ex."
		write-output '    $f1 = psexcel_getFormula $xls "Sheet1" "A1"'
		write-output '    $f2 = psexcel_getFormula $xls 2 "B2"'
		write-output ""
		return
	}

	return $objExcel.Worksheets.Item($sheet).Range($range).Formula
}

#
# psexcel_setFormula - Set formula on range on sheet
#
function psexcel_setFormula($objExcel, $sheet, $range, $formula) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_setFormula excel_object sheet range formula"
		write-output "Set formula on range on sheet."
		write-output "ex."
		write-output '    psexcel_setFormula $xls "Sheet1" "A1" "=A1+B2"'
		write-output '    psexcel_setFormula $xls 2 "B2" "=C3+D4"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).Formula = $formula
}

#
# psexcel_getBackgroundColor - Get index of background color from range on sheet
#
function psexcel_getBackgroundColor($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getBackgroundColor excel_object sheet range"
		write-output "Get index of background color from range on sheet."
		write-output "ex."
		write-output '    $colIndex1 = psexcel_getBackgroundColor $xls "Sheet1" "A1"'
		write-output '    $colIndex2 = psexcel_getBackgroundColor $xls 2 "B2"'
		write-output ""
		return
	}

	return $objExcel.Worksheets.Item($sheet).Range($range).interior.ColorIndex
}

#
# psexcel_setBackgroundColor - Set background color to color_index on range on sheet
#
function psexcel_setBackgroundColor($objExcel, $sheet, $range, $colorIndex) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_setBackgroundColor excel_object sheet range color_index"
		write-output "Set background color to color_index on range on sheet."
		write-output "ex."
		write-output '    psexcel_setBackgroundColor $xls "Sheet1" "A1" 1'
		write-output '    psexcel_setBackgroundColor $xls 2 "B2" 3'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).interior.ColorIndex = $colorIndex
}

#
# psexcel_getForegroundColor Get index of foreground color from range on sheet
#
function psexcel_getForegroundColor($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getForegroundColor excel_object sheet range"
		write-output "Get index of foreground color from range on sheet."
		write-output "ex."
		write-output '    $colIndex1 = psexcel_getForegroundColor $xls "Sheet1" "A1"'
		write-output '    $colIndex2 = psexcel_getForegroundColor $xls 2 "B2"'
		write-output ""
		return
	}

	return $objExcel.Worksheets.Item($sheet).Range($range).Font.ColorIndex
}

#
# psexcel_setForegroundColor - Set foreground color to color_index on range on sheet
#
function psexcel_setForegroundColor($objExcel, $sheet, $range, $colorIndex) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_setForegroundColor excel_object sheet range color_index"
		write-output "Set foreground color to color_index on range on sheet."
		write-output "ex."
		write-output '    psexcel_setForegroundColor $xls "Sheet1" "A1" 1'
		write-output '    psexcel_setForegroundColor $xls 2 "B2" 3'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).Font.ColorIndex = $colorIndex
}

#
# psexcel_getBold - Get $true or $false about bold from range on sheet
#
function psexcel_getBold($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getBold excel_object sheet range"
		write-output 'Get $true or $false about bold from range on sheet.'
		write-output "ex."
		write-output '    $boolean1 = psexcel_getBold $xls "Sheet1" "A1"'
		write-output '    $boolean2 = psexcel_getBold $xls 2 "B2"'
		write-output ""
		return
	}

	return $objExcel.Worksheets.Item($sheet).Range($range).Font.Bold
}

#
# psexcel_turnonBold - Turn on bold on range on sheet
#
function psexcel_turnonBold($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_turnonBold excel_object sheet range"
		write-output 'Turn on bold on range on sheet.'
		write-output "ex."
		write-output '    psexcel_turnonBold $xls "Sheet1" "A1"'
		write-output '    psexcel_turnonBold $xls 2 "B2"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).Font.Bold = $true
}

#
# psexcel_turnoffBold - Turn off bold on range on sheet
#
function psexcel_turnoffBold($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_turnoffBold excel_object sheet range"
		write-output 'Turn off bold on range on sheet.'
		write-output "ex."
		write-output '    psexcel_turnoffBold $xls "Sheet1" "A1"'
		write-output '    psexcel_turnoffBold $xls 2 "B2"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).Font.Bold = $false
}

#
# psexcel_toggleBold - Toggle bold on range on sheet
#
function psexcel_toggleBold($objExcel, $sheet, $range) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_toggleBold excel_object sheet range"
		write-output 'Toggle bold on range on sheet.'
		write-output "ex."
		write-output '    psexcel_toggleBold $xls "Sheet1" "A1"'
		write-output '    psexcel_toggleBold $xls 2 "B2"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Range($range).Font.Bold = -Not($objExcel.Worksheets.Item($sheet).Range($range).Font.Bold)
}

#
# psexcel_getSheetCount - Get count of sheets
#
function psexcel_getSheetCount($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getSheetCount excel_object"
		write-output 'Get count of sheets.'
		write-output "ex."
		write-output '    $sc = psexcel_getSheetCount $xls'
		write-output ""
		return
	}

	return $objExcel.Workbooks.item(1).Sheets.Count
}

#
# psexcel_getSheetName - Get name of sheet
#
function psexcel_getSheetName($objExcel, $sheet) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getSheetName excel_object sheet"
		write-output 'Get name of sheet.'
		write-output "ex."
		write-output '    $sn = psexcel_getSheetName $xls 2'
		write-output ""
		return
	}

	return $objExcel.Worksheets.Item($sheet).Name
}

#
# psexcel_setSheetName - Set name of sheet
#
function psexcel_setSheetName($objExcel, $sheet, $name) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_setSheetName excel_object sheet name"
		write-output 'Set name of sheet.'
		write-output "ex."
		write-output '    psexcel_setSheetName $xls "Sheet1" "SHEET1"'
		write-output '    psexcel_setSheetName $xls 2 "SHEET2"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).Name = $name
}

#
# psexcel_getActiveSheetName Get name of active sheet
#
function psexcel_getActiveSheetName($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_getActiveSheetName excel_object"
		write-output 'Get name of active sheet.'
		write-output "ex."
		write-output '    $sn = psexcel_getActiveSheetName $xls'
		write-output ""
		return
	}

	return $objExcel.Workbooks.item(1).ActiveSheet.Name
}

#
# psexcel_setActiveSheetName - Set name of active sheet
#
function psexcel_setActiveSheetName($objExcel, $name) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_setActiveSheetName excel_object name"
		write-output 'Set name of active sheet.'
		write-output "ex."
		write-output '    psexcel_setActiveSheetName $xls "SHEET1"'
		write-output ""
		return
	}

	for ($i = 1; $i -le $objExcel.Workbooks.item(1).Sheets.Count; $i++){
		if ($objExcel.Worksheets.Item($i).Name -eq $objExcel.Workbooks.item(1).ActiveSheet.Name){
			$objExcel.Worksheets.Item($i).Name = $name
			break
		}
	}
}

#
# psexcel_copyCell - Copy range of cell to another cell
#
function psexcel_copyCell($objExcel, $srcSheet, $srcRange, $destSheet, $destCell) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_copyCell excel_object source_sheet source_range dest_sheet dest_cell"
		write-output 'Copy range of cell to another cell.'
		write-output "ex."
		write-output '    psexcel_copyCell $xls "Sheet1" "A1:C3" "sheet2" "D4"'
		write-output '    psexcel_copyCell $xls 1 "A1:C3" 2 "D4"'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($srcSheet).Range($srcRange).copy($objExcel.Worksheets.Item($destSheet).Range($destCell)) | out-null
}

#
# psexcel_preview - Print preview sheet
#
function psexcel_preview($objExcel, $sheet){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_preview excel_object sheet"
		write-output 'Print preview sheet.'
		write-output "ex."
		write-output '    psexcel_preview $xls "Sheet1"'
		write-output '    psexcel_preview $xls 2'
		write-output ""
		return
	}

	$objExcel.Visible = $true
	$objExcel.Worksheets.Item($sheet).PrintPreview($true)
	$objExcel.Visible = $false
}

#
# psexcel_print - Print sheet
#
function psexcel_print($objExcel, $sheet){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_print excel_object sheet"
		write-output 'Print sheet.'
		write-output "ex."
		write-output '    psexcel_print $xls "Sheet1"'
		write-output '    psexcel_print $xls 2'
		write-output ""
		return
	}

	$objExcel.Worksheets.Item($sheet).PrintOut()
}

#
# psexcel_turnonVisible - Turn on visible
#
function psexcel_turnonVisible($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_turnonVisible excel_object"
		write-output "Turn on visible."
		write-output "ex."
		write-output '    psexcel_turnonVisible $xls'
		write-output ""
		return
	}
	$objExcel.Visible = $true
}

#
# psexcel_turnoffVisible - Turn off visible
#
function psexcel_turnoffVisible($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_turnoffVisible excel_object"
		write-output "Turn off visible."
		write-output "ex."
		write-output '    psexcel_turnoffVisible $xls'
		write-output ""
		return
	}
	$objExcel.Visible = $false
}

#
# psexcel_turnonAlert - Turn on displayAlerts
#
function psexcel_turnonAlert($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_turnonAlert excel_object"
		write-output "Turn on displayAlerts."
		write-output "ex."
		write-output '    psexcel_turnonAlert $xls'
		write-output ""
		return
	}
	$objExcel.DisplayAlerts = $true
}

#
# psexcel_turnoffAlert - Turn off displayAlerts
#
function psexcel_turnoffAlert($objExcel) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_turnoffAlert excel_object"
		write-output "Turn off displayAlerts."
		write-output "ex."
		write-output '    psexcel_turnoffAlert $xls'
		write-output ""
		return
	}
	$objExcel.DisplayAlerts = $false
}

#
# psprov - Print formatted data with overlay
#
Function psprov(){
	begin{
		Function psprov_prpv($rec) {
			foreach ($i in $formatArray.keys){
				psexcel_setCell $oOverlay 1 $i $rec.split($delimiter)[$formatArray[$i]]
			}
			if ($previewSw -eq $true){
				psexcel_preview $oOverlay 1
			}else{
				psexcel_print $oOverlay 1
			}
		}

		$helpSw = $false
		$previewSw = $false
		$delimiter = ","
		$overlay_path = ""
		$input_path = ""
		$format_path = ""
		$formatArray = @{}
		$files = @{}
		$filesIndex = 0

		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output 'Usage: psprov [-p] [-d "DELIMITER"] -o OVERLAY.XLS [-i INPUT.CSV] -f FORMAT.TXT'
				write-output 'Print formatted data with overlay.'
				write-output ''
				write-output '  -p                    preview mode.'
				write-output '  -d "DELIMITER"        use DELIMITER instead of comma for INPUT.CSV.'
				write-output '  -o OVERLAY.XLS        overlay definition by excel.'
				write-output '  -i INPUT.CSV          input data in csv format. If omitted then read stdin.'
				write-output '  -f FORMAT.TXT         format definition for INPUT.CSV.'
				write-output '                        each line should have like "A1=1"'
				write-output '                        "A1=1" means "A1" cell in OVERLAY.XLS should be setted to "1st column" in INPUT.CSV'
				write-output ''
				return
			}elseif ($args[$i] -eq "-p"){
				$previewSw = $true
			}elseif ($args[$i] -eq "-d"){
				$i++
				$delimiter = $args[$i]
			}elseif ($args[$i] -eq "-o"){
				$i++
				$overlay_path = $args[$i]
			}elseif ($args[$i] -eq "-i"){
				$i++
				$input_path = psabspath $args[$i]
			}elseif ($args[$i] -eq "-f"){
				$i++
				$format_path = $args[$i]
			}else{
#				$files[$filesIndex] = (resolve-path $args[$i]).Path
				$files[$filesIndex] = psabspath $args[$i]
				$filesIndex++
			}
		}

		if ($overlay_path -ne ""){
			$oOverlay = psexcel_open $overlay_path
		}

		if ($input_path -ne ""){
			$oIn = psopen -r $input_path
		}

		if ($format_path -ne ""){
			$fmt = psopen -r $format_path
			while (($rec = $fmt.readLine()) -ne $null){
				if ($rec.split("=")[0] -match "^[A-z][A-z]*[0-9][0-9]*$"){
					$formatArray[$rec.split("=")[0]] = [int]$rec.split("=")[1] - 1
				}
			}
			$fmt.close()
		}
	}
	process{
		if ($helpSw -eq $false){
			if ($input_path -ne ""){
				while (($rec = $oIn.readLine()) -ne $null){
					psprov_prpv $rec
				}
			}else{
				psprov_prpv $_
			}
		}
	}
	end{
		if ($helpSw -eq $false){
			if ($input_path -ne ""){
				$oIn.close()
			}
			if ($overlay_path -ne ""){
				psexcel_close $oOverlay
			}
		}
	}
}
