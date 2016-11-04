# tools4powershell

tools4powershell is a collection of some useful functions for powershell :)

## pscat - concatenate files and print on the standard output
```
Usage: pscat [-h|--help] [-n] [input ...]
Concatenate input(s), or standard input, to standard output.

  -n        number all output lines
```

## psgrep - print lines matching a pattern
```
Usage: psgrep [-h|--help] [-v] [-i] regex [input ...]
Search for regex in each input or standard input.

  -v        select non-matching lines
  -i        ignore case distinctions
```

## pswcl - print newline counts for each file
```
Usage: pswcl [-h|--help] [input ...]
Print newline counts for each input, and a total line if more than one input is specified.
```

## pshead - output the first part of files
```
Usage: pshead [-h|--help] [-l line_number] [input ...]
Print the first 10 lines of each input to standard output.
With no input, read standard input.

  -l line_number        print the first line_number lines instead of the first 10
```

## pstail - output the last part of files
```
Usage: pstail [-h|--help] [-l line_number] [input ...]
Print the last 10 lines of each input to standard output.
With no input, read standard input.

  -l line_number        print the last line_number lines instead of the last 10
```

## pscut - remove sections from each line of files
```
Usage: pscut [-h|--help] [-d "delimiter"] -i "index,..." [input ...]
Print selected parts of lines from each input to standard output.
With no input, read standard input.

  -d "delimiter"        use "delimiter" instead of "," for field delimiter
  -i "index,..."        select only these fields(0 origin)
```

## pstee - read from standard input and write to standard output and files
```
Usage: pstee [-h|--help] [output ...]
Copy standard input to each output, and also to standard output.
```

## psuniq - report or omit repeated lines
```
Usage: psuniq [-h|--help] [-d|-c] [input ...]
Filter adjacent matching lines from input (or standard input),
writing to standard output.
With no options, matching lines are merged to the first occurrence.

  -d        only print duplicate lines
  -c        prefix lines by the number of occurrences
```

## psjoin - join lines of two files on a common field
```
Usage: psjoin [-h|--help] [-d "delimiter"] [-1 "index,..."] [-2 "index,..."] [-a [m|1|2|12|21]] [-m [1|2]] input1 input2
For each pair of input lines with identical join fields, write a line to
standard output.  The default join field is the first, delimited by ",".

  -d "delimiter"        use "delimiter" as input and output field separator instead of ","
  -1 "index,..."        join on this index(s) of file 1 (default 0)
  -2 "index,..."        join on this index(s) of file 2 (default 0)
  -a m                    write only matching lines (default)
     1                    write only unpairable lines from input1
     2                    write only unpairable lines from input2
     12                   write all lines from input1 and matching lines from input2
     21                   write all lines from input2 and matching lines from input1
  -m [1|2]                specify input which has multiple join fields
```

## psxls2csv - convert excel to csv
```
Usage: psxls2csv [-h|--help] [-i input] [-s sheet] [-o [output|-]]
Convert specified excel sheet to csv.
If input is not specified, all excel files in current directory will be converted.
If output is not specified, input will be converted into same filename, but with extention ".csv".
If "-" is specified for "-o" option, input will be converted into stdout.

BUGS
  If input is not specified and output is specified, only last excel sheet in current directory will remain in output file.
```

## psprint - print arguments to standard output
```
Usage: psprint [-h|--help] [arg ...]
Print arguments to standard output.
```

## pstmpfile - get temporary file
```
Usage: pstmpfile [-h|--help]
Print temporary file path to standard output.
```

## psabspath - get absolute file path
```
Usage: psabspath [-h|--help] file
Print absolute file path for file to standard output.
```

## psreadpassword - get password safely from console
```
Usage: psreadpassword [-h|--help]
Get password safely from console and decrypt and print it to console.
```

## psconwrite - output arguments to console without newline
```
Usage: psconwrite [-h|--help] [arg ...]
Print arguments to console without newline.
```

## psconwriteline - output arguments to console with newline
```
Usage: psconwriteline [-h|--help] [arg ...]
Print arguments to console with newline.
```

## psconreadline - read from console
```
Usage: psconreadline [-h|--help]
Read line and print it to console with newline.
```

## psopen - Open IO.Stream
```
Usage: psopen [-h|--help] [[-r|-w|-a] [inputfile|output]] [-e encoding]
Open IO.Stream and get object.
```
