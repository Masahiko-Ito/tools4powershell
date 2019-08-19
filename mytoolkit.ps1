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
		$encoding1 = 0
		$encoding2 = 0
		for ($i = 0; $i -lt $args.length; $i++){
			if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
				$helpSw = $true
				write-output "Usage: psjoin [-h|--help] [-d ""delimiter""] [-1 ""index,...""] [-2 ""index,...""] [-a [m|1|2|12|21]] [-m [1|2]] [-e1 encoding] [-e2 encoding] input1 input2"
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
				write-output "  -e1 encoding            encoding for file 1(default 0 means Default)"
				write-output "  -e2 encoding            encoding for file 2(default 0 means Default)"
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
			}elseif ($args[$i] -eq "-e1"){
				$i++
				$encoding1 = $args[$i]
			}elseif ($args[$i] -eq "-e2"){
				$i++
				$encoding2 = $args[$i]
			}else{
#				$files[$filesIndex] = (resolve-path $args[$i]).Path
				$files[$filesIndex] = psabspath $args[$i]
				$filesIndex++
			}
		}
#		$oIn1 = New-Object System.IO.StreamReader($files[0],[Text.Encoding]::Default)
#		$oIn2 = New-Object System.IO.StreamReader($files[1],[Text.Encoding]::Default)
		if ($encoding1 -eq "utf8n" -or 
		    $encoding1 -eq "UTF8N" -or 
		    $encoding1 -eq "utf-8n" -or 
		    $encoding1 -eq "UTF-8N"){
			$enc1 = New-Object System.Text.UTF8Encoding $False
		}else{
			$enc1 = [Text.Encoding]::GetEncoding($encoding1)
		}
		$oIn1 = New-Object System.IO.StreamReader($files[0],$enc1)
		if ($encoding2 -eq "utf8n" -or 
		    $encoding2 -eq "UTF8N" -or 
		    $encoding2 -eq "utf-8n" -or 
		    $encoding2 -eq "UTF-8N"){
			$enc2 = New-Object System.Text.UTF8Encoding $False
		}else{
			$enc2 = [Text.Encoding]::GetEncoding($encoding1)
		}
		$oIn2 = New-Object System.IO.StreamReader($files[1],$enc2)
	}
	process{
		if ($helpSw -eq $false){
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
	
			$endSw = "off"
			$rec1 = $oIn1.readLine()
			$rec2 = $oIn2.readLine()
			$matchkey = $null
			while ($endSw -eq "off"){
				$key1 = mkkey $rec1 $keyidx1
				$key2 = mkkey $rec2 $keyidx2
				if ($action -eq "1"){
					if (islt $key1 $key2){
						if ($matchkey -eq $null -or $matchkey -cne $key1){
							write-output $rec1
						}
						$rec1 = $oIn1.readLine()
					}elseif (islt $key2 $key1){
						$rec2 = $oIn2.readLine()
					}else{
						$matchkey = $key1
						if ($multikey -eq 1){
							$rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$rec2 = $oIn2.readLine()
						}else{
							$rec1 = $oIn1.readLine()
							$rec2 = $oIn2.readLine()
						}
					}
				}elseif ($action -eq "2"){
					if (islt $key2 $key1){
						if ($matchkey -eq $null -or $matchkey -cne $key2){
							write-output $rec2
						}
						$rec2 = $oIn2.readLine()
					}elseif (islt $key1 $key2){
						$rec1 = $oIn1.readLine()
					}else{
						$matchkey = $key2
						if ($multikey -eq 1){
							$rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$rec2 = $oIn2.readLine()
						}else{
							$rec1 = $oIn1.readLine()
							$rec2 = $oIn2.readLine()
						}
					}
				}elseif ($action -eq "12"){
					if (islt $key1 $key2){
						if ($matchkey -eq $null -or $matchkey -cne $key1){
							write-output $rec1
						}
						$rec1 = $oIn1.readLine()
					}elseif (islt $key2 $key1){
						$rec2 = $oIn2.readLine()
					}else{
						$out = $rec1 + $delimiter + $rec2
						write-output $out
						$matchkey = $key1
						if ($multikey -eq 1){
							$rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$rec2 = $oIn2.readLine()
						}else{
							$rec1 = $oIn1.readLine()
							$rec2 = $oIn2.readLine()
						}
					}
				}elseif ($action -eq "21"){
					if (islt $key2 $key1){
						if ($matchkey -eq $null -or $matchkey -cne $key2){
							write-output $rec2
						}
						$rec2 = $oIn2.readLine()
					}elseif (islt $key1 $key2){
						$rec1 = $oIn1.readLine()
					}else{
						$out = $rec2 + $delimiter + $rec1
						write-output $out
						$matchkey = $key2
						if ($multikey -eq 1){
							$rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$rec2 = $oIn2.readLine()
						}else{
							$rec1 = $oIn1.readLine()
							$rec2 = $oIn2.readLine()
						}
					}
				}else{
					if (islt $key1 $key2){
						$rec1 = $oIn1.readLine()
					}elseif (islt $key2 $key1){
						$rec2 = $oIn2.readLine()
					}else{
						$out = $rec1 + $delimiter + $rec2
						write-output $out
						if ($multikey -eq 1){
							$rec1 = $oIn1.readLine()
						}elseif ($multikey -eq 2){
							$rec2 = $oIn2.readLine()
						}else{
							$rec1 = $oIn1.readLine()
							$rec2 = $oIn2.readLine()
						}
					}
				}
				if ($rec1 -eq $null -and $rec2 -eq $null){
					$endSw = "on"
				}
			}
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
#				if ($tabSw -eq "on"){
#					$TmpOutPath = [System.IO.Path]::GetTempFileName()
#					Get-Content $OutPath | Out-File -Encoding default $TmpOutPath 
#					Move-Item -Force -path $TmpOutPath -destination $OutPath
#				}
				if ($strOutput -eq "-"){
					get-content $OutPath
					remove-item $OutPath
				}
			}
		}else{
			if ($strInput -match '^[A-Za-z]:' -or $strInput -match '^\\'){
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
#			if ($tabSw -eq "on"){
#				$TmpOutPath = [System.IO.Path]::GetTempFileName()
#				Get-Content $OutPath | Out-File -Encoding default $TmpOutPath 
#				Move-Item -Force -path $TmpOutPath -destination $OutPath
#			}
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
# psxls2sheetname - get sheetname from excel book
#
function psxls2sheetname (){
	$input = ""
	$sheet = ""
	$helpSw = $false
	for ($i = 0; $i -lt $args.length; $i++){
		if ($args[$i] -eq "-h" -or $args[$i] -eq "--help"){
			$helpSw = $true
			write-output "Usage: psxls2sheetname [-h|--help] [-i input] [-s sheet]"
			write-output "Print sheet name identified by sheet."
			write-output "If sheet is omitted, print all sheet names."
			write-output ""
			return
		}elseif ($args[$i] -eq "-i"){
			$i++
			$input = $args[$i]
		}elseif ($args[$i] -eq "-s"){
			$i++
			$sheet = $args[$i]
		}
	}
	if ($helpSw -eq $false){
		$xls = psexcel_open $input
		if ($sheet -eq ""){
			for ($i = 1; $i -le (psexcel_getSheetCount $xls); $i++){
				psexcel_getSheetName $xls $i
			}
		}else{
			psexcel_getSheetName $xls $sheet
		}
		psexcel_close $xls
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
			if ($path -match '^[A-Za-z]:' -or $path -match '^\\'){
				$path
			}else{
#				(get-location).tostring() + "\" + $path
				(get-location).ProviderPath + "\" + $path
			}
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
# ex. Handle data as text
#     $inObj = psopen -r "input.txt"
#     $outObj = psopen -w "output.txt"
#     while (($rec = $inObj.readLine()) -ne $null){
#         $outObj.writeLine($rec)
#     }
#     $inObj.close()
#     $outObj.close()
#
# ex. Handle data as binary
#     $arrayByte = New-Object byte[] (1024)
#     $inObj = psopen -r "input.txt"
#     $outObj = psopen -w "output.txt"
#     $read_length = $inObj.BaseStream.Read($arrayByte, 0, $arrayByte.length)
#     while ($read_length > 0){
#         $outObj.BaseStream.Write($arrayByte, 0, $read_length)
#         $read_length = $inObj.BaseStream.Read($arrayByte, 0, $arrayByte.length)
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
				write-output "ex. Handle as text"
				write-output '  $inObj = psopen -r "input.txt"'
				write-output '  $outObj = psopen -w "output.txt"'
				write-output '  while (($rec = $inObj.readLine()) -ne $null){'
				write-output '      $outObj.writeLine($rec)'
				write-output '  }'
				write-output '  $inObj.close()'
				write-output '  $outObj.close()'
				write-output ""
				write-output "ex. Handle data as binary"
				write-output '  $arrayByte = New-Object byte[] (1024)'
				write-output '  $inObj = psopen -r "input.txt"'
				write-output '  $outObj = psopen -w "output.txt"'
				write-output '  $read_length = $inObj.BaseStream.Read($arrayByte, 0, $arrayByte.length)'
				write-output '  while ($read_length > 0){'
				write-output '      $outObj.BaseStream.Write($arrayByte, 0, $read_length)'
				write-output '      $read_length = $inObj.BaseStream.Read($arrayByte, 0, $arrayByte.length)'
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
					$inputfile = $inputfile -replace "^.*::", ""
				}
			}elseif ($args[$i] -eq "-w"){
				$iomode = "w"
				$i++
				if ($args[$i] -match "^[A-Za-z]:" -or $args[$i] -match "^\\"){
					$outputfile = $args[$i]
				}else{
					$outputfile = (get-location).tostring() + "\" + $args[$i]
					$outputfile = $outputfile -replace "^.*::", ""
				}
			}elseif ($args[$i] -eq "-a"){
				$iomode = "a"
				$i++
				if ($args[$i] -match "^[A-Za-z]:" -or $args[$i] -match "^\\"){
					$outputfile = $args[$i]
				}else{
					$outputfile = (get-location).tostring() + "\" + $args[$i]
					$outputfile = $outputfile -replace "^.*::", ""
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
			if ($encoding -eq "utf8n" -or 
			    $encoding -eq "UTF8N" -or 
			    $encoding -eq "utf-8n" -or 
			    $encoding -eq "UTF-8N"){
				$enc = New-Object System.Text.UTF8Encoding $False
			}else{
				$enc = [Text.Encoding]::GetEncoding($encoding)
			}
			if ($iomode -eq "r"){
				$objIO = New-Object System.IO.StreamReader($inputfile, $enc)
			}elseif ($iomode -eq "w"){
				$objIO = New-Object System.IO.StreamWriter($outputfile, $false, $enc)
			}else{
				$objIO = New-Object System.IO.StreamWriter($outputfile, $true, $enc)
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

	$objExcel.Workbooks.item(1).Save() | out-null
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
# psexcel_replaceString - Replace string to another string
#
function psexcel_replaceString($objExcel, $srcSheet, $srcRange, $beforeString, $afterString) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help" -or $objExcel -eq "-h" -or $objExcel -eq "--help"){
		write-output "Usage: psexcel_replaceString excel_object source_sheet source_range before_string after_string"
		write-output 'Replace string to another string.'
		write-output "ex."
		write-output '    psexcel_replaceString $xls "Sheet1" "A1:C3" "#1#" "Hoge"'
		write-output '    psexcel_replaceString $xls 1 "A1:C3" "#1#" "Hoge"'
		write-output ""
		return
	}

	$aPartOfString = 2
	$objExcel.Worksheets.Item($srcSheet).Range($srcRange).replace($beforeString, $afterString, $aPartOfString) | out-null
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

#
# psoracle_open - Connect to oracle
#
function psoracle_open($ip, $user, $password){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_open oracle_ip_address username password"
		write-output "Connect to oracle."
		write-output "ex."
		write-output '    $ocon = poracle_open "127.0.0.1" "taro" "himitsu"'
		write-output ""
		return
	}

	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")
	$ocon = New-Object System.Data.OracleClient.OracleConnection("Data Source=$ip;User ID=$user;Password=$password;Integrated Security=false;")
	$ocon.open()
	return $ocon
}

#
# psoracle_close - Disconnect from oracle
#
function psoracle_close($ocon){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_close oracle_connection"
		write-output "Disconnect from oracle."
		write-output "ex."
		write-output '    psoracle_close $ocon'
		write-output ""
		return
	}

	$ocon.Close()
	$ocon.Dispose()
}

#
# psoracle_createsql - Create SQL object with bind parameter
#
function psoracle_createsql($oc, $sql){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_createsql oracle_connection sql_command"
		write-output "Create SQL object with bind parameter."
		write-output "ex."
		write-output '    $ocmd = psoracle_createsql $ocon "select * from table where id=:id_value"'
		write-output '    ... something to do ...'
		write-output '    psoracle_free $ocmd'
		write-output ""
		return
	}

	$ocmd = New-Object System.Data.OracleClient.OracleCommand
	$ocmd.Connection = $oc
	$ocmd.CommandText = $sql

#	$otran = $oc.BeginTransaction()
#	$otran.Commit()
#	$ocmd.Transaction = $otran

	return $ocmd
}

#
# psoracle_bindsql - Bind real value to bind variable
#
function psoracle_bindsql($ocmd, $name, $value) {
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_bindsql oracle_command name value"
		write-output "Bind real value to bind variable."
		write-output "ex."
		write-output '    $id_value = psoracle_bindsql $ocmd "id_value" "12345"'
		write-output ""
		return
	}

	$bind = New-Object System.Data.OracleClient.OracleParameter($name, $value)
	[void]$ocmd.Parameters.Add($bind)
	return $bind
}

#
# psoracle_unbindsql - Unbind bind variable
#
function psoracle_unbindsql($ocmd, $bind){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_unbindsql oracle_command bind_object"
		write-output "Unbind bind variable."
		write-output "ex."
		write-output '    psoracle_unbindsql $ocmd $id_value'
		write-output ""
		return
	}

	[void]$ocmd.Parameters.Remove($bind)
}

#
# psoracle_execupdatesql - Exuceute SQL with update
#
function psoracle_execupdatesql($ocmd){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_execupdatesql oracle_command"
		write-output "Exuceute SQL with update."
		write-output "ex."
		write-output '    $count = psoracle_execupdatesql $ocmd'
		write-output '    if ($count -lt 0){'
		write-output '        write-output "Error: update SQL"'
		write-output '    }'
		write-output ""
		return
	}

	$effected_row = $ocmd.ExecuteNonQuery()
	return $effected_row
}

#
# psoracle_execsql - Exuceute SQL without update
#
function psoracle_execsql($ocmd){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_execsql oracle_command"
		write-output "Exuceute SQL without update."
		write-output "ex."
		write-output '    $ocr = psoracle_execsql $ocmd'
		write-output '    ... something to do ...'
		write-output '    psoracle_free $ocr'
		write-output ""
		return
	}

	$ord = $ocmd.ExecuteReader()
	return ,$ord
}

#
# psoracle_fetch - Fetch row
#
function psoracle_fetch([Ref]$ocr){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_fetch ([Ref]oracle_reader)"
		write-output "Fetch row."
		write-output "ex."
		write-output '    while (psoracle_fetch ([Ref]$ocr)){'
		write-output '        write-output $ocr["id"]'
		write-output '    }'
		write-output ""
		return
	}

	($ocr.value).Read()
}

#
# psoracle_free - Free object for oracle access
#
function psoracle_free($obj){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_free oracle_object"
		write-output "Free object for oracle access."
		write-output "ex."
		write-output '    psoracle_free $ocr'
		write-output '    psoracle_free $ocmd'
		write-output ""
		return
	}

	$obj.Dispose()
}

#
# psoracle_begin - Begin transaction
#
function psoracle_begin($ocon){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_begin oracle_connection"
		write-output "Begin transaction."
		write-output "ex."
		write-output '    $otran = psoracle_begin $ocon'
		write-output ""
		return
	}

	$otran = $ocon.BeginTransaction()
	return $otran
}

#
# psoracle_settran - Set transaction for oracle command
#
function psoracle_settran($ocmd, $otran){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_settran oracle_command oracle_transaction"
		write-output "Set transaction for oracle command."
		write-output "ex."
		write-output '    psoracle_settran $ocmd $otran'
		write-output ""
		return
	}

	$ocmd.Transaction = $otran
}

#
# psoracle_rollback - rollback transaction
#
function psoracle_rollback($otran){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_rollback oracle_transaction"
		write-output "Rollback transaction."
		write-output "ex."
		write-output '    psoracle_rollback $otran'
		write-output ""
		return
	}

	$otran.Rollback()
}

#
# psoracle_commit - commit transaction
#
function psoracle_commit($otran){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psoracle_commit oracle_transaction"
		write-output "Commit transaction."
		write-output "ex."
		write-output '    psoracle_commit $otran'
		write-output ""
		return
	}

	$otran.Commit()
}

#
# pssock_open - Open socket for client
#
Function pssock_open($addr, $port, $encoding = 0){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_open server_ip server_port [encoding]"
		write-output "Open socket for client."
		write-output "ex."
		write-output '    $param = pssock_open "127.0.0.1" "12345" "UTF-8"'
		write-output ""
		return
	}

	$client = New-Object System.Net.Sockets.TcpClient ($addr, $port)
	$stream = $client.GetStream()
	if ($encoding -eq "utf8n" -or 
	    $encoding -eq "UTF8N" -or 
	    $encoding -eq "utf-8n" -or 
	    $encoding -eq "UTF-8N"){
		$enc = New-Object System.Text.UTF8Encoding $False
	}else{
		$enc = [Text.Encoding]::GetEncoding($encoding)
	}
#	$reader = New-Object IO.StreamReader($stream,[Text.Encoding]::Default)
#	$writer = New-Object IO.StreamWriter($stream,[Text.Encoding]::Default)
	$reader = New-Object IO.StreamReader($stream,$enc)
	$writer = New-Object IO.StreamWriter($stream,$enc)
	$param = @{"client" = $client; "stream" = $stream; "writer" = $writer; "reader" = $reader}
	return $param
}

#
# pssock_close - Close socket for client
#
Function pssock_close($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_close socket_param"
		write-output "Close socket for client."
		write-output "ex."
		write-output '    pssock_close $param'
		write-output ""
		return
	}

	$param["writer"].Close()
	$param["reader"].Close()
	$param["stream"].Close()
	$param["client"].Close()
}

#
# pssock_readline - Read a line from socket
#
Function pssock_readline($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_readline socket_param"
		write-output "Read a line from socket."
		write-output "ex."
		write-output '    $line = pssock_readline $param'
		write-output ""
		return
	}

	try {
		$line = $param["reader"].readLine()
	}catch{
		$line = $null
	}
	return $line
}

#
# pssock_writeline - Write a line to socket
#
Function pssock_writeline($param, $line){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_writeline socket_param string"
		write-output "Write a line to socket."
		write-output "ex."
		write-output '    $stat = pssock_writeline $param $line'
		write-output ""
		return
	}

	$stat = $true
	try {
		$param["writer"].writeLine($line)
		$param["writer"].Flush()
	}catch{
		$stat = $false
	}
	return $stat
}

#
# pssock_read - Read data as binary from socket
#
Function pssock_read($param, $bufBytes, $length){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_read socket_param array_byte length"
		write-output "Read data as binary from socket."
		write-output "ex."
		write-output '    $arrayByte = New-Object byte[] (1024)'
		write-output '    $read_length = pssock_read $param $arrayByte $arrayByte.length'
		write-output ""
		return
	}

	try {
		$read_len = $param["reader"].BaseStream.Read($bufBytes, 0, $length)
	}catch{
		$read_len = -1
	}
	return $read_len
}

#
# pssock_write - Write data as binary to socket
#
Function pssock_write($param, $bufBytes, $length){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_write socket_param array_byte length"
		write-output "Write data as binary to socket."
		write-output "ex."
		write-output '    $arrayByte = New-Object byte[] (1024)'
		write-output '    $stat = pssock_write $param $arrayByte $arrayByte.length'
		write-output ""
		return
	}

	$stat = $true
	try {
		$param["writer"].BaseStream.Write($bufBytes, 0, $length)
		$param["writer"].BaseStream.Flush()
	}catch{
		$stat = $false
	}
	return $stat
}

#
# pssock_start - Start server
#
Function pssock_start($port){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_start server_port"
		write-output "Start server."
		write-output "ex."
		write-output '    $server = pssock_start "12345"'
		write-output ""
		return
	}

	$endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, $port)
	$server = New-Object System.Net.Sockets.TcpListener $endpoint
	$server.start()
	return $server
}

#
# pssock_stop - Stop server
#
Function pssock_stop($server){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_stop server"
		write-output "Stop server."
		write-output "ex."
		write-output '    pssock_stop $server'
		write-output ""
		return
	}

	$server.stop()
}

#
# pssock_accept - Accept connection from client
#
Function pssock_accept($server, $encoding = 0){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_accept server [encoding]"
		write-output "Accept connection from client."
		write-output "ex."
		write-output '    $param = pssock_accept $server "UTF-8"'
		write-output ""
		return
	}

	$client = $server.AcceptTcpClient()
	$stream = $client.GetStream()
	if ($encoding -eq "utf8n" -or 
	    $encoding -eq "UTF8N" -or 
	    $encoding -eq "utf-8n" -or 
	    $encoding -eq "UTF-8N"){
		$enc = New-Object System.Text.UTF8Encoding $False
	}else{
		$enc = [Text.Encoding]::GetEncoding($encoding)
	}
#	$reader = New-Object IO.StreamReader($stream,[Text.Encoding]::Default)
#	$writer = New-Object IO.StreamWriter($stream,[Text.Encoding]::Default)
	$reader = New-Object IO.StreamReader($stream,$enc)
	$writer = New-Object IO.StreamWriter($stream,$enc)
	$param = @{"client" = $client; "stream" = $stream; "reader" = $reader; "writer" = $writer}
	return $param
}

#
# pssock_unaccept - Unaccept(disconnect) connection from client
#
Function pssock_unaccept($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_unaccept socket_param"
		write-output "Unaccept(disconnect) connection from client."
		write-output "ex."
		write-output '    pssock_unaccept $param'
		write-output ""
		return
	}

	$param["writer"].Close()
	$param["reader"].Close()
	$param["stream"].Close()
	$param["client"].Close()
}

#
# pssock_getip - Get ip-address of client
#
Function pssock_getip($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_getip socket_param"
		write-output "Get ip-address of client."
		write-output "ex."
		write-output '    $ip = pssock_getip $param'
		write-output ""
		return
	}

	return $param["client"].client.RemoteEndPoint.Address
}

#
# pssock_getipstr - Get ip-address string of client
#
Function pssock_getipstr($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_getipstr socket_param"
		write-output "Get ip-address string of client."
		write-output "ex."
		write-output '    $ip = pssock_getipstr $param'
		write-output ""
		return
	}

	return $param["client"].client.RemoteEndPoint.Address.toString()
}

#
# pssock_getport - Get port of client
#
Function pssock_getport($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_getport socket_param"
		write-output "Get port of client."
		write-output "ex."
		write-output '    $ip = pssock_getport $param'
		write-output ""
		return
	}

	return $param["client"].client.RemoteEndPoint.Port
}

#
# pssock_getportstr - Get port string of client
#
Function pssock_getportstr($param){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: pssock_getportstr socket_param"
		write-output "Get port string of client."
		write-output "ex."
		write-output '    $ip = pssock_getportstr $param'
		write-output ""
		return
	}

	return $param["client"].client.RemoteEndPoint.Port.toString()
}

#
# psrunspc_getarraylist - Get System.Collections.ArrayList
#
Function psrunspc_getarraylist(){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_getarraylist"
		write-output "Get System.Collections.ArrayList."
		write-output "ex."
		write-output '    $refArray = psrunspc_getarraylist'
		write-output '    $Array0 = $ref.Array.value[0]'
		write-output ""
		return
	}
	$al = New-Object System.Collections.ArrayList
	return [Ref]$al
}

#
# psrunspc_open - Create and Open RunSpacePool
#
Function psrunspc_open($max_runspace){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_open max_runspace"
		write-output "Create and Open RunSpacePool."
		write-output "ex."
		write-output '    $rsp = psrunspc_open 10'
		write-output '    ... something to do ...'
		write-output '    psrunspc_close $rsp'
		write-output ""
		return
	}
	$rsp = [RunspaceFactory]::CreateRunspacePool(1, $max_runspace)
	$rsp.Open()
	return $rsp
}

#
# psrunspc_close - Close RunSpacePool
#
Function psrunspc_close($rsp){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_close run_space_pool"
		write-output "Close RunSpacePool."
		write-output "ex."
		write-output '    $rsp = psrunspc_open 10'
		write-output '    ... something to do ...'
		write-output '    psrunspc_close $rsp'
		write-output ""
		return
	}
	$rsp.Close()
}

#
# psrunspc_createthread - Create thread of powershell and add script to it
#
Function psrunspc_createthread($rsp, $script){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_createthread run_space_pool script_block"
		write-output "Create thread of powershell and add script to it."
		write-output "ex."
		write-output '    $ps = psrunspc_createthread $rsp $script'
		write-output ""
		return
	}
	$ps = [PowerShell]::Create()
	$ps.RunspacePool = $rsp
	$ps.AddScript($script) | out-null
	return $ps
}

#
# psrunspc_addargument - Add argument to thread of powershell
#
Function psrunspc_addargument($ps, $arg){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_addargument thread_of_powershellargument"
		write-output "Add argument to thread of powershell."
		write-output "ex."
		write-output '    psrunspc_addargument $argument'
		write-output ""
		return
	}
	$ps.AddArgument($arg) | out-null
}

#
# psrunspc_begin - Begin script in thread
#
Function psrunspc_begin($ps, $aryps, $arych){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_begin thread_of_powershell arraylist_of_powershell arraylist_of_child"
		write-output "Begin script in thread."
		write-output "ex."
		write-output '    $refAryps = psrunspc_getarraylist'
		write-output '    $refArych = psrunspc_getarraylist'
		write-output '    ... something to do ...'
		write-output '    psrunspc_begin $ps $refAryps.value $refArych.value'
		write-output ""
		return
	}
	$ch = $ps.BeginInvoke()
	$aryps.Add($ps) | Out-Null
	$arych.Add($ch) | Out-Null
}

#
# psrunspc_wait - Wait terminate of all child thread
#
Function psrunspc_wait($aryps, $arych){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_wait arraylist_of_powershell arraylist_of_child"
		write-output "Wait terminate of all child thread."
		write-output "ex."
		write-output '    psrunspc_wait $refAryps.value $refArych.value'
		write-output ""
		return
	}
	for ($i = 0; $i -lt $aryps.Count; $i++){
		$ps = $aryps[$i]
		$ch = $arych[$i]
		$result = $ps.EndInvoke($ch)
		$ps.Dispose()
		$aryps.removeAt($i)
		$arych.removeAt($i)
	}
}

#
# psrunspc_waitasync - Wait asynchronously terminate of all child thread
#
Function psrunspc_waitasync($aryps, $arych){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrunspc_waitasync arraylist_of_powershell arraylist_of_child"
		write-output "Wait asynchronously terminate of all child thread."
		write-output "ex."
		write-output '    psrunspc_waitasync $refAryps.value $refArych.value'
		write-output ""
		return
	}
	for ($i = 0; $i -lt $aryps.Count; $i++){
		$ps = $aryps[$i]
		$ch = $arych[$i]
		if ($ch.isCompleted){
			$result = $ps.EndInvoke($ch)
			$ps.Dispose()
			$aryps.removeAt($i)
			$arych.removeAt($i)
		}
	}
}

#
# psrpa_init - Initialize rpa environment
#
function psrpa_init{
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_init"
		write-output "Initialize rpa environment and get rpa_object."
		write-output "ex."
		write-output '    $rpa = psrpa_init'
		write-output ""
		return
	}
#	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#	[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	add-type -assemblyname System.Drawing
	add-type -assemblyname System.Windows.Forms
	add-type -assemblyname microsoft.visualbasic
	add-type -assemblyname system.windows.forms

	$source = @" 
		using System;
		using System.Runtime.InteropServices;
		using System.Windows.Automation;
		using System.Drawing;
		using System.Drawing.Imaging;
		public class Psrpa {
			[DllImport("user32.dll")]
			[return: MarshalAs(UnmanagedType.Bool)]
			public static extern bool MoveWindow(IntPtr hWnd,int X, int Y, int nWidth, int nHeight, bool bRepaint);
			[DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
			public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
			[StructLayout(LayoutKind.Sequential)]
			public struct MOUSEINPUT {
				public int dx;
				public int dy;
				public int mouseData;
				public int dwFlags;
				public int time;
				public int dwExtraInfo;
			};
			[StructLayout(LayoutKind.Sequential)]
			public struct KEYBDINPUT {
				public short wVk;
				public short wScan;
				public int dwFlags;
				public int time;
				public int dwExtraInfo;
			};
			[StructLayout(LayoutKind.Sequential)]
			public struct HARDWAREINPUT {
				public int uMsg;
				public short wParamL;
				public short wParamH;
			};
			[StructLayout(LayoutKind.Explicit)]
			public struct INPUT {
				[FieldOffset(0)]
				public int type;
				[FieldOffset(4)]
				public MOUSEINPUT no;
				[FieldOffset(4)]
				public KEYBDINPUT ki;
				[FieldOffset(4)]
				public HARDWAREINPUT hi;
			};

			[DllImport("user32.dll")]
			public extern static void SendInput(int nInputs, ref INPUT pInputs, int cbsize);
			[DllImport("user32.dll", EntryPoint = "MapVirtualKeyA")]
			public extern static int MapVirtualKey(int wCode, int wMapType);
			public const int INPUT_KEYBOARD = 1;
			public const int KEYEVENTF_KEYDOWN = 0x0;
			public const int KEYEVENTF_KEYUP = 0x2;
			public const int KEYEVENTF_EXTENDEDKEY = 0x1;
			// public void Send(Keys key, bool isEXTEND){
			public static void Send(short key, bool isDown, bool isEXTEND){
				INPUT inp = new INPUT();

				inp.type = INPUT_KEYBOARD;
				inp.ki.wVk = (short)key;
				inp.ki.wScan = (short)MapVirtualKey(inp.ki.wVk, 0);
				inp.ki.time = 0;
				inp.ki.dwExtraInfo = 0;
				if (isDown){
					inp.ki.dwFlags = ((isEXTEND)?(KEYEVENTF_EXTENDEDKEY):0x0)|KEYEVENTF_KEYDOWN;
				}else{
					inp.ki.dwFlags = ((isEXTEND)?(KEYEVENTF_EXTENDEDKEY):0x0)|KEYEVENTF_KEYUP;
				}
				SendInput(1, ref inp, Marshal.SizeOf(inp));
			}
			public static AutomationElement GetRootWindow(){ 
				return AutomationElement.RootElement; 
			}
			public static AutomationElement GetMainWindowByTitle(string title) {
				PropertyCondition cond = new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty, title);
				return AutomationElement.RootElement.FindFirst(TreeScope.Element | TreeScope.Children, cond);
			}
			[System.Runtime.InteropServices.DllImport("msvcrt.dll",CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
			private static extern int memcmp(byte[] b1, byte[] b2, UIntPtr count);
			public static bool CompareImage(Bitmap img1, Bitmap img2){
				ImageConverter ic = new ImageConverter();
				byte[] byte1 = (byte[])ic.ConvertTo(img1, typeof(byte[]));
				byte[] byte2 = (byte[])ic.ConvertTo(img2, typeof(byte[]));
				/* Console.WriteLine("{0} {1}", byte1.Length, byte2.Length); */
				if (byte1.Length != byte2.Length){
					return false;
				}
				return CompareByte(byte1, byte2);
			}
			public static bool CompareByte(byte[] byte1, byte[] byte2){
				if (byte1.Length != byte2.Length){
					return false;
				}
				return memcmp(byte1, byte2, new UIntPtr((uint)byte1.Length)) == 0;
			}
			public static int[] SearchImage(Bitmap scrimg, Bitmap fileimg){
				BitmapData scrimgdata = scrimg.LockBits(
					new Rectangle(0, 0, scrimg.Width, scrimg.Height),
					ImageLockMode.ReadWrite,
					PixelFormat.Format32bppArgb
				);
				BitmapData fileimgdata = fileimg.LockBits(
					new Rectangle(0, 0, fileimg.Width, fileimg.Height),
					ImageLockMode.ReadWrite,
					PixelFormat.Format32bppArgb
				);
				int[] pos = new int[4] {-1, -1, -1, -1};
				byte[] scrimgbyte = new byte[scrimg.Width * scrimg.Height * 4];
				byte[] fileimgbyte = new byte[fileimg.Width * fileimg.Height * 4];
				Marshal.Copy(scrimgdata.Scan0, scrimgbyte, 0, scrimgbyte.Length);
				Marshal.Copy(fileimgdata.Scan0, fileimgbyte, 0, fileimgbyte.Length);
				byte[] scrimggraybyte = new byte[scrimg.Width * scrimg.Height];
				byte[] fileimggraybyte = new byte[fileimg.Width * fileimg.Height];
				int pixsize, j, x1, y1, x2, y2;
				bool isFound;

				pixsize = scrimg.Width * scrimg.Height * 4;
				j = 0;
				for (int i = 0; i < pixsize; i += 4){
					scrimggraybyte[j] = (byte)((scrimgbyte[i] + scrimgbyte[i + 1] + scrimgbyte[i + 2] + scrimgbyte[i + 3]) >> 2);
					j++;
				}
				pixsize = fileimg.Width * fileimg.Height * 4;
				j = 0;
				for (int i = 0; i < pixsize; i += 4){
					fileimggraybyte[j] = (byte)((fileimgbyte[i] + fileimgbyte[i + 1] + fileimgbyte[i + 2] + fileimgbyte[i + 3]) >> 2);
					j++;
				}
				x1 = 0;
				y1 = 0;
				x2 = scrimg.Width - fileimg.Width;
				y2 = scrimg.Height - fileimg.Height;
				isFound = false;
				for (int y = y1; !isFound && y <= y2; y++){
					for (int x = x1; !isFound && x <= x2; x++){
						isFound = true;
						for (int fy = 0; isFound && fy < fileimg.Height; fy++){
							int dy = y + fy;
							int swy = scrimg.Width * dy;
							int fwy = fileimg.Width * fy;
							for (int fx = 0; isFound && fx < fileimg.Width; fx++){
								int dx = x + fx;
								if (scrimggraybyte[swy + dx] != fileimggraybyte[fwy + fx]){
									isFound = false;
								}
							}
						}
						if (isFound){
							pos[0] = x;
							pos[1] = y;
							pos[2] = x + fileimg.Width;
							pos[3] = y + fileimg.Height;
						}
					}
				}
				return pos;
			}
		} 
"@ 
	Add-Type -Language CSharp -TypeDefinition $source -ReferencedAssemblies("UIAutomationClient", "UIAutomationTypes", "System.Drawing")

	$param = @{"BeforeWait" = 300;
		"AfterWait" = 300
	}
	return $param
}

#
# psrpa_setBeforeWait - Set a time(ms) for waiting before functions(psrpa_*)
#
function psrpa_setBeforeWait($rpa, $timems){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_setBeforeWait rpa_object wait_time_ms"
		write-output 'Set a time(ms) for waiting before functions(psrpa_*).'
		write-output ""
		return
	}
	$rpa["BeforeWait"] = $timems
}

#
# psrpa_setAfterWait - Set a time(ms) for waiting after functions(psrpa_*)
#
function psrpa_setAfterWait($rpa, $timems){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_setAfterWait rpa_object wait_time_ms"
		write-output 'Set a time(ms) for waiting after functions(psrpa_*).'
		write-output ""
		return
	}
	$rpa["AfterWait"] = $timems
}

#
# psrpa_showMouse - Show current mouse position for debug purpose
#
function psrpa_showMouse($rpa){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_showMouse rpa_object"
		write-output "Show current mouse position for debug purpose."
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	while ($true){
		$x = [System.Windows.Forms.Cursor]::Position.X
		$y = [System.Windows.Forms.Cursor]::Position.Y
		write-output "$x $y"
		sleep -Second 1
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_showMouseByClick - Show current mouse position for psrpa_setMouse and psrpa_clickPoint by click
#
function psrpa_showMouseByClick($rpa, $wait = 5){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_showMouseByClick rpa_object [wait_sec]"
		write-output "Show current mouse position for psrpa_setMouse and psrpa_clickPoint by click."
		write-output "Press any key to terminate."
		write-output "    wait_sec    default 5"
		write-output "ex."
		write-output '    $wait_sec = 10'
		write-output '    psrpa_showMouseByClick $rpa $wait_sec'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	Start-Sleep $wait

	$pwidth = (
		gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentHorizontalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$pheight = (
		gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentVerticalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$img = New-Object System.Drawing.Bitmap([int]$pwidth, [int]$pheight)
	$gr = [System.Drawing.Graphics]::FromImage($img)
	$gr.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
	$gr.CopyFromScreen(
		(New-Object System.Drawing.Point(0,0)), 
		(New-Object System.Drawing.Point(0,0)), 
		$img.Size
	)
	$lwidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
	$lheight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
	$w = New-Object System.Windows.Forms.Form
	$w.Text = "psrpa_showMouseByClick"
	$w.keyPreview = $true
	$w.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	$pb = New-Object System.Windows.Forms.PictureBox
	$w.ClientSize = $img.Size
	$pb.ClientSize = $img.Size
	$pb.Image = $img
	$w.Controls.Add($pb)

	$w_keydown = {
		$w.Close()
		$img.Dispose()
		$gr.Dispose()
		$pb.Dispose()
	}
	$w.Add_KeyDown($w_keydown)

	$pb_click = {
		$x = [System.Windows.Forms.Cursor]::Position.X
		$y = [System.Windows.Forms.Cursor]::Position.Y
		if ($_.Button -eq "Left"){
			$button = "left"
		}elseif ($_.Button -eq "Middle"){
			$button = "middle"
		}elseif ($_.Button -eq "Right"){
			$button = "right"
		}else{
			$button = "left"
		}
		write-host ('psrpa_setMouse $rpa ' + "$x $y")
		write-host ('psrpa_clickPoint $rpa ' + "$x $y" + ' "' + $button + '" "click"')
		write-host ('')
	}
	$pb.Add_Click($pb_click)

	$pb.Cursor = [System.Windows.Forms.Cursors]::Cross
	$stat = $w.ShowDialog()

	Start-Sleep -Milliseconds $rpa["AfterWait"]
}


#
# psrpa_setMouse - Set mouse position
#
function psrpa_setMouse($rpa, $x, $y){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_setMouse rpa_object x_position y_position"
		write-output "Set mouse position."
		write-output "ex."
		write-output '    $x = 10'
		write-output '    $y = 20'
		write-output '    psrpa_setMouse $rpa $x $y'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	[system.windows.forms.cursor]::position = new-object system.drawing.point($x,$y)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_click - Click mouse button
#
function psrpa_click($rpa, $button, $action){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_click rpa_object mouse_button click_action"
		write-output "Click mouse button."
		write-output "    mouse_button left, LEFT, l, L"
		write-output "                 middle, MIDDLE, m, M"
		write-output "                 right, RIGHT, r, R"
		write-output "    click_action click, CLICK, 1"
		write-output "                 2click, 2CLICK, 2"
		write-output "                 3click, 3CLICK, 3"
		write-output "                 down, DOWN, d, D"
		write-output "                 up, UP, u, U"
		write-output "ex."
		write-output '    psrpa_click $rpa "left" "click"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	if ($button -eq "left" -or 
		$button -eq "LEFT" -or
		$button -eq "l" -or
		$button -eq "L"){
		$down = 0x0000002
		$up = 0x0000004
	}elseif ($button -eq "middle" -or 
		$button -eq "MIDDLE" -or
		$button -eq "m" -or
		$button -eq "M"){
		$down = 0x0000020
		$up = 0x0000040
	}elseif ($button -eq "right" -or 
		$button -eq "RIGHT" -or
		$button -eq "r" -or
		$button -eq "R"){
		$down = 0x0000008
		$up = 0x0000010
	}else{
		$down = 0x0000002
		$up = 0x0000004
	}

	if ($action -eq "click" -or
		$action -eq "CLICK" -or
		$action -eq "1"){
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
	}elseif ($action -eq "2click" -or
		$action -eq "2CLICK" -or
		$action -eq "2"){
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
	}elseif ($action -eq "3click" -or
		$action -eq "3CLICK" -or
		$action -eq "3"){
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
	}elseif ($action -eq "down" -or
		$action -eq "DOWN" -or
		$action -eq "d" -or
		$action -eq "D"){
		[Psrpa]::mouse_event($down,0,0,0,0)
	}elseif ($action -eq "up" -or
		$action -eq "UP" -or
		$action -eq "u" -or
		$action -eq "U"){
		[Psrpa]::mouse_event($up,0,0,0,0)
	}else{
		[Psrpa]::mouse_event($down,0,0,0,0)
		[Psrpa]::mouse_event($up,0,0,0,0)
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_clickPoint - Set mouse position and click mouse button
#
function psrpa_clickPoint($rpa, $x, $y, $button, $action){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_clickPoint rpa_object x_position y_position mouse_button click_action"
		write-output "Set mouse position and click mouse button."
		write-output "    mouse_button left, LEFT, l, L"
		write-output "                 middle, MIDDLE, m, M"
		write-output "                 right, RIGHT, r, R"
		write-output "    click_action click, CLICK, 1"
		write-output "                 2click, 2CLICK, 2"
		write-output "                 3click, 3CLICK, 3"
		write-output "                 down, DOWN, d, D"
		write-output "                 up, UP, u, U"
		write-output "ex."
		write-output '    $x = 10'
		write-output '    $y = 20'
		write-output '    psrpa_clickPoint $rpa $x $y "left" "click"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	psrpa_setMouse $rpa $x $y
	psrpa_click $rpa $button $action
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_showAppTitle - Show application and title
#
function psrpa_showAppTitle($rpa){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_showAppTitle rpa_object"
		write-output "Show application and title."
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$ps = get-process
	foreach ($process in $ps){
		write-output ('"' + $process.Name + '" "' + $process.MainWindowTitle + '"')
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_activateWindow - Activate specified window
#
function psrpa_activateWindow ($rpa, $application, $title){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_activateWindow rpa_object application title"
		write-output "Activate specified window"
		write-output '    application    "" means "^$", $null means "^.*$"'
		write-output '    title          "" means "^$", $null means "^.*$"'
		write-output "ex."
		write-output '    $application = "notepad"'
		write-output '    $title = "no title - memo pad"'
		write-output '    psrpa_activateWindow $rpa $application $title'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	if ($application -eq $null){
		$application = "^.*$"
	}elseif ($application -eq ""){
		$application = "^$"
	}
	if ($title -eq $null){
		$title = "^.*$"
	}elseif ($title -eq ""){
		$title = "^$"
	}
	$ps = get-process | where-object {$_.Name -match $application}
	foreach ($process in $ps){
		if ($process.MainWindowTitle -ne ""){
			if ($process.MainWindowTitle -match $title){
				[Microsoft.Visualbasic.Interaction]::AppActivate($process.ID)
			}
		}
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_setWindow - Set window position and size
#
#function psrpa_setWindow($rpa, $application, $title, $x, $y, $width, $height){
function psrpa_setWindow{
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_setWindow rpa_object application title x_position y_position width height"
		write-output "Set window position and size"
		write-output '    application    "" means "^$", $null means "^.*$"'
		write-output '    title          "" means "^$", $null means "^.*$"'
		write-output "ex."
		write-output '    $application = "notepad"'
		write-output '    $title = "no title - memo pad"'
		write-output '    $x = 10'
		write-output '    $y = 20'
		write-output '    $width = 300'
		write-output '    $height = 400'
		write-output '    psrpa_setWindow $rpa $application $title $x $y $width $height'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$rpa = $args[0]
	$application = $args[1]
	$title = $args[2]
	$x = $args[3]
	$y = $args[4]
	$width = $args[5]
	$height = $args[6]
	if ($application -eq $null){
		$application = "^.*$"
	}elseif ($application -eq ""){
		$application = "^$"
	}
	if ($title -eq $null){
		$title = "^.*$"
	}elseif ($title -eq ""){
		$title = "^$"
	}
	$ps = get-process | where-object {$_.Name -match $application}
	foreach ($process in $ps){
		if ($process.MainWindowTitle -ne ""){
			if ($process.MainWindowTitle -match $title){
				[Psrpa]::MoveWindow($process.MainWindowHandle, $x, $y, $width, $height, $true) | out-null
			}
		}
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_sendKeys - Send keys(string, function, special keys etc)
#
function psrpa_sendKeys ($rpa, $string){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_sendKeys rpa_object string"
		write-output "Send keys(string, function, special keys etc)"
		write-output "    string any-string"
		write-output "           {BS} for backspace"
		write-output "           {DEL} for delete"
		write-output "           {PRTSC} for print screen"
		write-output "           {ENTER} for enter"
		write-output "           {LEFT} for left arrow"
		write-output "           {F1} for F1"
		write-output "           ^a for CTRL-a"
		write-output "           +a for SHIFT-a"
		write-output "           %a for ALT-a"
		write-output "           other keys - search by google :P"
		write-output "ex."
		write-output '    psrpa_sendKeys $rpa "any string"'
		write-output '    psrpa_sendKeys $rpa "{BS}"'
		write-output '    psrpa_sendKeys $rpa "^a"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	[system.windows.forms.sendkeys]::sendwait($string)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}


#
# psrpa_sendKeyEx - Send a key(Both of normal key and extended key are acceptable)
#
function psrpa_sendKeyEx ($rpa, $virtual_keycode, $action, $isExtended){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_sendKeyEx rpa_object virtual_keycode action isExtended"
		write-output "Send a key(Both of normal key and extended key are acceptable)"
		write-output '    virtual_keycode see https://msdn.microsoft.com/ja-jp/windows/desktop/dd375731'
		write-output '    action "down", "DOWN", "d", "D"'
		write-output '           "up", "UP", "u", "U"'
		write-output '           "send", "SEND", "downup", "DOWNUP"'
		write-output '    isExtended $true or $false'
		write-output '               see https://docs.microsoft.com/en-us/windows/win32/inputdev/about-keyboard-input'
		write-output '                   https://kazhat.at.webry.info/201809/article_4.html'
		write-output "ex."
		write-output '    $right_ctrl = 0xa3'
		write-output '    psrpa_sendKeyEx $rpa $right_ctrl "send" $true'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	if ($action -eq "down" -or
		$action -eq "DOWN" -or
		$action -eq "d" -or
		$action -eq "D"){
		[Psrpa]::Send($virtual_keycode, $true, $isExtended)
	}elseif ($action -eq "up" -or
		$action -eq "UP" -or
		$action -eq "u" -or
		$action -eq "U"){
		[Psrpa]::Send($virtual_keycode, $false, $isExtended)
	}elseif ($action -eq "send" -or
		$action -eq "SEND" -or
		$action -eq "downup" -or
		$action -eq "DOWNUP"){
		[Psrpa]::Send($virtual_keycode, $true, $isExtended)
		Start-Sleep -Milliseconds 100
		[Psrpa]::Send($virtual_keycode, $false, $isExtended)
	}else{
		[Psrpa]::Send($virtual_keycode, $true, $isExtended)
		Start-Sleep -Milliseconds 100
		[Psrpa]::Send($virtual_keycode, $false, $isExtended)
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}
#
# psrpa_setClipboard - Set clipboard to string
#
function psrpa_setClipboard($rpa, $string){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_setClipboard rpa_object string"
		write-output "Set clipboard to string."
		write-output "Caution!"
		write-output "    powershell.exe have to be invoked with -sta option."
		write-output "ex."
		write-output '    psrpa_setClipboard $rpa "any string"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	[Windows.Forms.Clipboard]::SetText($string)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_getClipboard - Get string from clipboard
#
function psrpa_getClipboard($rpa){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_getClipboard rpa_object"
		write-output "Get string from clipboard."
		write-output "Caution!"
		write-output "    powershell.exe have to be invoked with -sta option."
		write-output "ex."
		write-output '    $str = psrpa_getClipboard $rpa'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return [Windows.Forms.Clipboard]::GetText()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_getBmp - Get image into file(BMP) 
#
function psrpa_getBmp($rpa, $x1, $y1, $x2, $y2, $bmpfile){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_getBmp rpa_object left_x top_y right_x bottom_y output.bmp"
		write-output "Get image into file(BMP)."
		write-output "Output file name will be (left_x)_(top_y)_(right_x)_(bottom_y)_output.bmp."
		write-output "ex."
		write-output '    psrpa_getBmp $rpa 10 10 200 100 "icon.bmp"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$outputfile = "$x1" + "_" + "$y1" + "_" + "$x2" + "_" + "$y2" + "_" + $bmpfile
	$dstimg = psrpa_getBmpFromInnerFunction $rpa $x1 $y1 $x2 $y2
	$dstimg.Save((psabspath $outputfile), [System.Drawing.Imaging.ImageFormat]::Bmp)
	$dstimg.Dispose()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_getBmpByClick - Get image into file(BMP) by click
#
function psrpa_getBmpByClick($rpa, $bmpfile, $wait = 5){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_getBmpByClick rpa_object output.bmp [wait_sec]"
		write-output "Get image into file(BMP) by click."
		write-output "Output file name will be (left_x)_(top_y)_(right_x)_(bottom_y)_output.bmp."
		write-output "Mouse cursor is CROSS for pointing (left_x, top_y)."
		write-output "                HAND  for pointing (right_x, bottom_y)."
		write-output "Press any key to terminate."
		write-output "    wait_sec    default 5"
		write-output "ex."
		write-output '    $wait_sec = 10'
		write-output '    psrpa_getBmpByClick $rpa "icon.bmp" $wait_sec'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	Start-Sleep $wait

	$pwidth = (
		gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentHorizontalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$pheight = (
		gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentVerticalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$img = New-Object System.Drawing.Bitmap([int]$pwidth, [int]$pheight)
	$gr = [System.Drawing.Graphics]::FromImage($img)
	$gr.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
	$gr.CopyFromScreen(
		(New-Object System.Drawing.Point(0,0)), 
		(New-Object System.Drawing.Point(0,0)), 
		$img.Size
	)
	$lwidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
	$lheight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
	$w = New-Object System.Windows.Forms.Form
	$w.Text = "psrpa_getBmpByClick"
	$w.keyPreview = $true
	$w.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
	$pb = New-Object System.Windows.Forms.PictureBox
	$w.ClientSize = $img.Size
	$pb.ClientSize = $img.Size
	$pb.Image = $img
	$w.Controls.Add($pb)

	$w_keydown = {
		$w.Close()
		$img.Dispose()
		$gr.Dispose()
		$pb.Dispose()
	}
	$w.Add_KeyDown($w_keydown)

	$script:psrpa_getBmpByClick_click_seq = 0
	$pb_click = {
		if ($script:psrpa_getBmpByClick_click_seq -eq 0){
			$script:psrpa_getBmpByClick_x1 = [System.Windows.Forms.Cursor]::Position.X
			$script:psrpa_getBmpByClick_y1 = [System.Windows.Forms.Cursor]::Position.Y
			$script:psrpa_getBmpByClick_click_seq = 1
			$pb.Cursor = [System.Windows.Forms.Cursors]::Hand
		}else{
			$x1 = $script:psrpa_getBmpByClick_x1
			$y1 = $script:psrpa_getBmpByClick_y1
			$x2 = [System.Windows.Forms.Cursor]::Position.X
			$y2 = [System.Windows.Forms.Cursor]::Position.Y

			$width = $x2 - $x1
			$height = $y2 - $y1
			$rect = New-Object System.Drawing.Rectangle($x1, $y1, $width, $height)
			$dstimg = $img.Clone($rect, $img.PixelFormat)
			$outputfile = "$x1" + "_" + "$y1" + "_" + "$x2" + "_" + "$y2" + "_" + $bmpfile
			$dstimg.Save((psabspath $outputfile), [System.Drawing.Imaging.ImageFormat]::Bmp)
			write-host ('psrpa_compareBmp $rpa ' + "$x1 $y1 $x2 $y2 $outputfile")
			$dstimg.Dispose()
			$rect = $null
			$script:psrpa_getBmpByClick_click_seq = 0
			$pb.Cursor = [System.Windows.Forms.Cursors]::Cross
		}
	}
	$pb.Add_Click($pb_click)

	$pb.Cursor = [System.Windows.Forms.Cursors]::Cross
	$stat = $w.ShowDialog()

	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_compareBmp - Compare specifoed rectangle and bmpfile 
#
function psrpa_compareBmp($rpa, $x1, $y1, $x2, $y2, $bmpfile){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_compareBmp rpa_object left_x top_x right_x bottom_y input.bmp"
		write-output "Compare specifoed rectangle and bmpfile."
		write-output 'Return $true when specified rectangle and input.bmp are same.'
		write-output "ex."
		write-output '    $bool = psrpa_compareBmp $rpa 10 10 200 100 "icon.bmp"'
		write-output '    if ($bool){'
		write-output '        write-output "match"'
		write-output '    }else{'
		write-output '        write-output "not match"'
		write-output '    }'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$dstimg = psrpa_getBmpFromInnerFunction $rpa $x1 $y1 $x2 $y2
	$fileimg = New-Object System.Drawing.Bitmap((psabspath $bmpfile))

	$img1 = New-Object System.Drawing.Bitmap($dstimg.Size.Width, $dstimg.Size.Height)
	$gr1 = [System.Drawing.Graphics]::FromImage($img1)
	$gr1.DrawImage($dstimg, 0, 0, $dstimg.Size.Width, $dstimg.Size.Height)
	$img2 = New-Object System.Drawing.Bitmap($dstimg.Size.Width, $dstimg.Size.Height)
	$gr2 = [System.Drawing.Graphics]::FromImage($img2)
	$gr2.DrawImage($fileimg, 0, 0, $dstimg.Size.Width, $dstimg.Size.Height)

	$isSame = [Psrpa]::CompareImage($img1, $img2)

	$dstimg.Dispose()
	$fileimg.Dispose()
	$img1.Dispose()
	$img2.Dispose()
	$gr1.Dispose()
	$gr2.Dispose()

	return $isSame
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_searchBmp - Search bmpfile in screen
#
function psrpa_searchBmp($rpa, $x1, $y1, $x2, $y2, $bmpfile){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_searchBmp rpa_object left_x top_x right_x bottom_y input.bmp"
		write-output "Search bmpfile in screen."
		write-output 'Return @(left,top,right,bottom) when bmpfile is found in screen.'
		write-output "ex."
		write-output '    $pos = psrpa_searchBmp $rpa $null $null $null $null "icon.bmp"'
		write-output '    if ($pos[0] < 0){'
		write-output '        write-output "not found"'
		write-output '    }else{'
		write-output '        write-output "found"'
		write-output '        $left = $pos[0]'
		write-output '        $top = $pos[1]'
		write-output '        $right = $pos[2]'
		write-output '        $bottom = $pos[3]'
		write-output '    }'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$pwidth = (gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentHorizontalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$pheight = (gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentVerticalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	if ($x1 -eq $null -or $x1 -eq ""){
		$x1 = 0
	}
	if ($y1 -eq $null -or $y1 -eq ""){
		$y1 = 0
	}
	if ($x2 -eq $null -or $x2 -eq ""){
		$x2 = $pwidth
	}
	if ($y2 -eq $null -or $y2 -eq ""){
		$y2 = $pheight
	}
	$scrimg = psrpa_getBmpFromInnerFunction $rpa 0 0 $pwidth $pheight
	$fileimg = New-Object System.Drawing.Bitmap((psabspath $bmpfile))
	$pos_array = [Psrpa]::SearchImage($scrimg, $fileimg)
	$scrimg.Dispose()
	$fileimg.Dispose()
	return $pos_array
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# Called from psrpa_getBmp and psrpa_compareBmp
#
function psrpa_getBmpFromInnerFunction($rpa, $x1, $y1, $x2, $y2){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Sorry, internal usage only."
		return
	}
###	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$width = $x2 - $x1
	$height = $y2 - $y1
	$pwidth = (
		gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentHorizontalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$pheight = (
		gwmi win32_videocontroller | 
		out-string -stream | 
		select-string "CurrentVerticalResolution" | 
		foreach{$_ -replace "^.*: *",""} | 
		sort | 
		select-object -Last 1
	)
	$img = New-Object System.Drawing.Bitmap([int]$pwidth, [int]$pheight)
	$gr = [System.Drawing.Graphics]::FromImage($img)
	$gr.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
	$gr.CopyFromScreen(
		(New-Object System.Drawing.Point(0,0)), 
		(New-Object System.Drawing.Point(0,0)), 
		$img.Size
	)
	$rect = New-Object System.Drawing.Rectangle($x1, $y1, $width, $height)
	$dstimg = $img.Clone($rect, $img.PixelFormat)
	$img.Dispose()
	$gr.Dispose()
	$rect = $null
	return $dstimg
###	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_show - Show all ui-automation element information
#
function psrpa_uia_show($rpa){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_show rpa_object"
		write-output "Show all ui-automation element."
		write-output "ex."
		write-output '    psrpa_uia_show $rpa'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	[Psrpa]::GetRootWindow().FindAll(
		[System.Windows.Automation.TreeScope]::Children,
		[System.Windows.Automation.Condition]::TrueCondition) |
	%{

		"========================================================================="
		$_.FindAll(
			[System.Windows.Automation.TreeScope]::SubTree,
			[System.Windows.Automation.Condition]::TrueCondition) |
		%{
			"ClassName = " + $_.Current.ClassName
#			"ControlType.Id = " + $_.Current.ControlType.Id.tostring()
			"LocalizedControlType = " + $_.Current.LocalizedControlType
			"Name = " + $_.Current.Name
#			"ProcessId = " + $_.Current.ProcessId.tostring()
			$_.GetSupportedPatterns() |
			%{
				"SupportedPattern = " + $_.Id.tostring() + ":" + $_.ProgrammaticName.tostring()
			}
			"-------------------------------------------------------------------------"
		}
	}
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_get - Get ui-automation element
#
function psrpa_uia_get($rpa, $element, $classname, $localizedcontroltype, $name){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_get rpa_object parent_element class_name localized_controlname name"
		write-output "Get ui-automation element."
		write-output '    parent_element        "" or $null mean root-window'
		write-output '    class_name            "" means "^$", $null means "^.*$"'
		write-output '    localized_controlname'
		write-output '    name'
		write-output "ex."
		write-output '    $app = psrpa_uia_get $rpa $null "Notepad" "ウィンドウ" "無題 - メモ帳"'
		write-output '    $help = psrpa_uia_get $rpa $app "" "メニュー項目" "ヘルプ\(H\)"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	if ($element -eq $null -or $element -eq ""){
		$element = [Psrpa]::GetRootWindow()
	}
	if ($classname -eq $null){
		$classname = "^.*$"
	}elseif ($classname -eq ""){
		$classname = "^$"
	}
	if ($localizedcontroltype -eq $null){
		$localizedcontroltype = "^.*$"
	}elseif ($localizedcontroltype -eq ""){
		$localizedcontroltype = "^$"
	}
	if ($name -eq $null){
		$name = "^.*$"
	}elseif ($name -eq ""){
		$name = "^$"
	}
	return ($element.FindAll(
			[System.Windows.Automation.TreeScope]::SubTree,
			[System.Windows.Automation.Condition]::TrueCondition) |
		where-object{$_.Current.ClassName -match $classname -and 
			$_.Current.LocalizedControlType -match $localizedcontroltype -and
			$_.Current.Name -match $name}
		)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGeometry - Get geometry of ui-automation element(*Pattern)
#
function psrpa_uia_getGeometry($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGeometry rpa_object element"
		write-output "Get geometry of ui-automation element(*Pattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $geo = psrpa_uia_getGeometry $rpa $elm'
		write-output '    $x = $geo[0]'
		write-output '    $y = $geo[1]'
		write-output '    $width = $geo[2]'
		write-output '    $height = $geo[3]'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$x = $element.Current.BoundingRectangle.X
	$y = $element.Current.BoundingRectangle.Y
	$width = $element.Current.BoundingRectangle.Width
	$height = $element.Current.BoundingRectangle.Height
	return @($x, $y, $width, $height)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_invoke - Invoke ui-automation element(InvokePattern)
#
function psrpa_uia_invoke($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_invoke rpa_object element"
		write-output "Invoke ui-automation element(InvokePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_invoke $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_expand - Expand ui-automation element(ExpandCollapsePattern)
#
function psrpa_uia_expand($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_expand rpa_object element"
		write-output "Expand ui-automation element(ExpandCollapsePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_expand $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_close - Close ui-automation element(WindowPattern)
#
function psrpa_uia_close($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_close rpa_object element"
		write-output "Close ui-automation element(WindowPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_close $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_setValue - Set value into ui-automation element(ValuePattern)
#
function psrpa_uia_setValue($rpa, $element, $value){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_setValue rpa_object element value"
		write-output "Set value into ui-automation element(ValuePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_setValue $rpa $elm "any-string"'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($value)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getText - Get text from ui-automation element(TextPattern)
#
function psrpa_uia_getText($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getText rpa_object element"
		write-output "Get text from ui-automation element(TextPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $text = psrpa_uia_getText $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern).DocumentRange.getText(65535)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_toggle - Toggle ui-automation element(TogglePattern)
#
function psrpa_uia_toggle($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_toggle rpa_object element"
		write-output "Toggle ui-automation element(TogglePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_toggle $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_select - Select ui-automation element(SelectionItem)
#
function psrpa_uia_select($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_select rpa_object element"
		write-output "Select ui-automation element(SelectionItem)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_select $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_move - Move ui-automation element(TransformPattern)
#
function psrpa_uia_move($rpa, $element, $x, $y){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_move rpa_object element x y"
		write-output "Move ui-automation element(TransformPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $x = 10'
		write-output '    $y = 20'
		write-output '    psrpa_uia_move $rpa $elm $x $y'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.TransformPattern]::Pattern).Move($x, $y)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_resize - Resize ui-automation element(TransformPattern)
#
#function psrpa_uia_resize($rpa, $element, $width, $height){
function psrpa_uia_resize{
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_resize rpa_object element width height"
		write-output "Resize ui-automation element(TransformPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $width = 100'
		write-output '    $height = 200'
		write-output '    psrpa_uia_resize $rpa $elm $width $height'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$rpa = $args[0]
	$element = $args[1]
	$width = $args[2]
	$height = $args[3]
	$element.GetCurrentPattern([System.Windows.Automation.TransformPattern]::Pattern).Resize($width, $height)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridClassName - Get grid classname of ui-automation element(GridPattern)
#
function psrpa_uia_getGridClassName($rpa, $element, $row, $col){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridClassName rpa_object element row col"
		write-output "Get grid classname of ui-automation element(GridPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $row = 0'
		write-output '    $col = 0'
		write-output '    $classname = psrpa_uia_getGridClassName $rpa $elm $row $col'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern).GetItem($row, $col).current.ClassName
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridControlType - Get grid controltype of ui-automation element(GridPattern)
#
function psrpa_uia_getGridControlType($rpa, $element, $row, $col){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridControlType rpa_object element row col"
		write-output "Get grid controltype of ui-automation element(GridPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $row = 0'
		write-output '    $col = 0'
		write-output '    $controltype = psrpa_uia_getGridControlType $rpa $elm $row $col'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern).GetItem($row, $col).current.LocalizedControlType
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridName - Get grid name of ui-automation element(GridPattern)
#
function psrpa_uia_getGridName($rpa, $element, $row, $col){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridName rpa_object element row col"
		write-output "Get grid name of ui-automation element(GridPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $row = 0'
		write-output '    $col = 0'
		write-output '    $name = psrpa_uia_getGridName $rpa $elm $row $col'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern).GetItem($row, $col).current.Name
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridColumn - Get grid colmun of ui-automation element(GridPattern)
#
function psrpa_uia_getGridColumn($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridColumn rpa_object element"
		write-output "Get grid column of ui-automation element(GridPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $col = psrpa_uia_getGridColumn $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern).Current.ColumnCount
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridRow - Get grid row of ui-automation element(GridPattern)
#
function psrpa_uia_getGridRow($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridRow rpa_object element"
		write-output "Get grid row of ui-automation element(GridPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $row = psrpa_uia_getGridRow $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern).Current.RowCount
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridItemColumn - Get griditem column of ui-automation element(GridItemPattern)
#
function psrpa_uia_getGridItemColumn($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridItemColumn rpa_object element"
		write-output "Get griditem column of ui-automation element(GridItemPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $column = psrpa_uia_getGridItemColumn $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridItemPattern]::Pattern).Current.Column
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridItemColumnSpan - Get griditem column span of ui-automation element(GridItemPattern)
#
function psrpa_uia_getGridItemColumnSpan($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridItemColumnSpan rpa_object element"
		write-output "Get griditem column span of ui-automation element(GridItemPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $columnspan = psrpa_uia_getGridItemColumnSpan $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridItemPattern]::Pattern).Current.ColumnSpan
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridItemRow - Get griditem row of ui-automation element(GridItemPattern)
#
function psrpa_uia_getGridItemRow($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridItemRow rpa_object element"
		write-output "Get griditem row of ui-automation element(GridItemPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $row = psrpa_uia_getGridItemRow $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridItemPattern]::Pattern).Current.Row
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getGridItemRowSpan - Get griditem row span of ui-automation element(GridItemPattern)
#
function psrpa_uia_getGridItemRowSpan($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getGridItemRowSpan rpa_object element"
		write-output "Get griditem row span of ui-automation element(GridItemPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $rowspan = psrpa_uia_getGridItemRowSpan $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.GridItemPattern]::Pattern).Current.RowSpan
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_setRangeValue - Set RangeValue of ui-automation element(RangeValuePattern)
#
function psrpa_uia_setRangeValue($rpa, $element, $value){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_setRangeValue rpa_object element value"
		write-output "Set RangeValue of ui-automation element(RangeValuePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_setRangeValue $rpa $elm 10'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern).SetValue($value)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getRangeValue - Get RangeValue of ui-automation element(RangeValuePattern)
#
function psrpa_uia_getRangeValue($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getRangeValue rpa_object element"
		write-output "Get RangeValue of ui-automation element(RangeValuePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $value = psrpa_uia_getRangeValue $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern).Current.Value
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getRangeValueMax - Get Maximum RangeValue of ui-automation element(RangeValuePattern)
#
function psrpa_uia_getRangeValueMax($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getRangeValueMax rpa_object element"
		write-output "Get Maximum RangeValue of ui-automation element(RangeValuePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $max = psrpa_uia_getRangeValueMax $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern).Current.Maximum
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getRangeValueMin - Get Minimum RangeValue of ui-automation element(RangeValuePattern)
#
function psrpa_uia_getRangeValueMin($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getRangeValueMin rpa_object element"
		write-output "Get Minimum RangeValue of ui-automation element(RangeValuePattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $min = psrpa_uia_getRangeValueMin $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern).Current.Minimum
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_setScroll - Set Scroll of ui-automation element(ScrollPattern)
#
function psrpa_uia_setScroll($rpa, $element, $x_percent, $y_percent){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_setScroll rpa_object element horizontal_percent vertical_percent"
		write-output "Set Scroll of ui-automation element(ScrollPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    psrpa_uia_setScroll $rpa $elm 10 100'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$element.GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern).SetScrollPercent($x_percent, $y_percent)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getScroll - Get Scroll of ui-automation element(ScrollPattern)
#
function psrpa_uia_getScroll($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getScroll rpa_object element"
		write-output "Get Scroll of ui-automation element(ScrollPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $scroll = psrpa_uia_getScroll $rpa $elm'
		write-output '    $scroll_horizontal = $scroll[0]'
		write-output '    $scroll_vertical = $scroll[1]'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	$hp = $element.GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern).Current.HorizontalScrollPercent
	$vp = $element.GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern).Current.VerticalScrollPercent
	return @($hp, $vp)
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getScrollHview - Get Horizontal view size(%) of Scroll of ui-automation element(ScrollPattern)
#
function psrpa_uia_getScrollHview($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getScrollHview rpa_object element"
		write-output "Get Horizontal view size(%) of Scroll of ui-automation element(ScrollPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $horizontal_view_percent = psrpa_uia_getScrollHview $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern).Current.HorizontalViewSize
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getScrollVview - Get Vertical view size(%) of Scroll of ui-automation element(ScrollPattern)
#
function psrpa_uia_getScrollVview($rpa, $element){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getScrollVview rpa_object element"
		write-output "Get Vertical view size(%) of Scroll of ui-automation element(ScrollPattern)."
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $vertical_view_percent = psrpa_uia_getScrollVview $rpa $elm'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	return $element.GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern).Current.VerticalViewSize
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}

#
# psrpa_uia_getPattern - Get pattern from ui-automation element
#
function psrpa_uia_getPattern($rpa, $element, $pattern){
	if ($args[0] -eq "-h" -or $args[0] -eq "--help"){
		write-output "Usage: psrpa_uia_getPattern rpa_object element pattern"
		write-output "Get pattern from ui-automation element."
		write-output '    pattern "Invoke"'
		write-output '            "ExpandCollapse"'
		write-output '            "Window"'
		write-output '            "Value"'
		write-output '            "Text"'
		write-output '            "Toggle"'
		write-output '            "SelectionItem"'
		write-output '            "Grid"'
		write-output '            "GridItem"'
		write-output '            "Selection"'
		write-output '            "RangeValue"'
		write-output '            "Scroll"'
		write-output '            "MultipleView"'
		write-output '            "Table"'
		write-output '            "TableItem"'
		write-output '            "ScrollItem"'
		write-output "ex."
		write-output '    $elm = psrpa_uia_get ...snip...'
		write-output '    $pattern = psrpa_uia_getPattern $rpa $elm "Invoke"'
		write-output '    $pattern.Invoke()'
		write-output ""
		return
	}
	Start-Sleep -Milliseconds $rpa["BeforeWait"]
	if ($pattern -eq "Invoke"){
		$ret = element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
	}elseif ($pattern -eq "ExpandCollapse"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
	}elseif ($pattern -eq "Window"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
	}elseif ($pattern -eq "Value"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
	}elseif ($pattern -eq "Text"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
	}elseif ($pattern -eq "Toggle"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
	}elseif ($pattern -eq "SelectionItem"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
	}elseif ($pattern -eq "Grid"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern)
	}elseif ($pattern -eq "GridItem"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.GridItemPattern]::Pattern)
	}elseif ($pattern -eq "RangeValue"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.RangeValuePattern]::Pattern)
	}elseif ($pattern -eq "Scroll"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.ScrollPattern]::Pattern)
#
	}elseif ($pattern -eq "Selection"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern)
	}elseif ($pattern -eq "MultipleView"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.MultipleViewPattern]::Pattern)
	}elseif ($pattern -eq "Table"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.TablePattern]::Pattern)
	}elseif ($pattern -eq "TableItem"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.TableItemPattern]::Pattern)
	}elseif ($pattern -eq "ScrollItem"){
		$ret = $element.GetCurrentPattern([System.Windows.Automation.ScrollItemPattern]::Pattern)
	}
	return $ret
	Start-Sleep -Milliseconds $rpa["AfterWait"]
}
