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
Usage: psjoin [-h|--help] [-d "delimiter"] [-1 "index,..."] [-2 "index,..."] [-a [m|1|2|12|21]] [-m [1|2]] [-e1 encoding] [-e2 encoding] input1 input2
For each pair of input lines with identical join fields, write a line to
standard output.  The default join field is the first, delimited by ",".

  -d "delimiter"        use "delimiter" as input and output field separator instead of ","
  -1 "index,..."        join on this index(s) of file 1 (default 0)
  -2 "index,..."        join on this index(s) of file 2 (default 0)
  -a m                  write only matching lines (default)
     1                  write only unpairable lines from input1
     2                  write only unpairable lines from input2
     12                 write all lines from input1 and matching lines from input2
     21                 write all lines from input2 and matching lines from input1
  -m [1|2]              specify input which has multiple join fields
  -e1 encoding          encoding for file 1(default 0 means Default)
  -e2 encoding          encoding for file 2(default 0 means Default)
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

## psexcel_open - Open excel_file and get excel_object for it
```
Usage: psexcel_open excel_file
Open excel_file and get excel_object for it.
```

## psexcel_update - Save by overwrite
```
Usage: psexcel_update excel_object
Save by overwrite.
```

## psexcel_save - Save into excel_file
```
Usage: psexcel_save excel_object excel_file
Save into excel_file.
```

## psexcel_close - Close excel file by excel_object
```
Usage: psexcel_close excel_object
Close excel file by excel_object.
```

## psexcel_getCell - Get value from range on sheet
```
Usage: psexcel_getCell excel_object sheet range
Get value from range on sheet.
```

## psexcel_setCell Set value on range on sheet
```
Usage: psexcel_setCell excel_object sheet range value
Set value on range on sheet.
```

## psexcel_getFormula - Get formula from range on sheet
```
Usage: psexcel_getFormula excel_object sheet range
Get formula from range on sheet.
```

## psexcel_setFormula - Set formula on range on sheet
```
Usage: psexcel_setFormula excel_object sheet range formula
Set formula on range on sheet.
```

## psexcel_getBackgroundColor - Get index of background color from range on sheet
```
Usage: psexcel_getBackgroundColor excel_object sheet range
Get index of background color from range on sheet.
```

## psexcel_setBackgroundColor - Set background color to color_index on range on sheet
```
Usage: psexcel_setBackgroundColor excel_object sheet range color_index
Set background color to color_index on range on sheet.
```

## psexcel_getForegroundColor Get index of foreground color from range on sheet
```
Usage: psexcel_getForegroundColor excel_object sheet range
Get index of foreground color from range on sheet.
```

## psexcel_setForegroundColor - Set foreground color to color_index on range on sheet
```
Usage: psexcel_setForegroundColor excel_object sheet range color_index
Set foreground color to color_index on range on sheet.
```

## psexcel_getBold - Get $true or $false about bold from range on sheet
```
Usage: psexcel_getBold excel_object sheet range
Get $true or $false about bold from range on sheet.
```

## psexcel_turnonBold - Turn on bold on range on sheet
```
Usage: psexcel_turnonBold excel_object sheet range
Turn on bold on range on sheet.
```

## psexcel_turnoffBold - Turn off bold on range on sheet
```
Usage: psexcel_turnoffBold excel_object sheet range
Turn off bold on range on sheet.
```

## psexcel_toggleBold - Toggle bold on range on sheet
```
Usage: psexcel_toggleBold excel_object sheet range
Toggle bold on range on sheet.
```

## psexcel_getSheetCount - Get count of sheets
```
Usage: psexcel_getSheetCount excel_object
Get count of sheets.
```

## psexcel_getSheetName - Get name of sheet
```
Usage: psexcel_getSheetName excel_object sheet
Get name of sheet.
```

## psexcel_setSheetName - Set name of sheet
```
Usage: psexcel_setSheetName excel_object sheet name
Set name of sheet.
```

## psexcel_getActiveSheetName Get name of active sheet
```
Usage: psexcel_getActiveSheetName excel_object
Get name of active sheet.
```

## psexcel_setActiveSheetName - Set name of active sheet
```
Usage: psexcel_setActiveSheetName excel_object name
Set name of active sheet.
```

## psexcel_copyCell - Copy range of cell to another cell
```
Usage: psexcel_copyCell excel_object source_sheet source_range dest_sheet dest_cell
Copy range of cell to another cell.
```

## psexcel_replaceString - Replace string to another string
```
Usage: psexcel_replaceString excel_object sheet range before_string after_string
Replace string to another string.
```

## psexcel_preview - Print preview sheet
```
Usage: psexcel_preview excel_object sheet
Print preview sheet.
```

## psexcel_print - Print sheet
```
Usage: psexcel_print excel_object sheet
Print sheet.
```

## psexcel_turnonVisible - Turn on visible
```
Usage: psexcel_turnonVisible excel_object
Turn on visible.
```

## psexcel_turnoffVisible - Turn off visible
```
Usage: psexcel_turnoffVisible excel_object
Turn off visible.
```

## psexcel_turnonAlert - Turn on displayAlerts
```
Usage: psexcel_turnonAlert excel_object
Turn on displayAlerts.
```

## psexcel_turnoffAlert - Turn off displayAlerts
```
Usage: psexcel_turnoffAlert excel_object
Turn off displayAlerts.
```

## psprov - Print formatted data with overlay
```
Usage: psprov [-p] [-d "DELIMITER"] -o OVERLAY.XLS [-i INPUT.CSV] -f FORMAT.TXT
Print formatted data with overlay

  -p                    preview mode.
  -d "DELIMITER"        use DELIMITER instead of comma for INPUT.CSV.
  -o OVERLAY.XLS        overlay definition by excel.
  -i INPUT.CSV          input data in csv format. If omitted then read stdin.
  -f FORMAT.TXT         format definition for INPUT.CSV.
                        each line should have like "A1=1"
                        "A1=1" means "A1" cell in OVERLAY.XLS should be setted to "1st column" in INPUT.CSV
```

## psoracle_open - Connect to oracle
```
Usage: psoracle_open oracle_ip_address username password
Connect to oracle.
```

## psoracle_close - Disconnect from oracle
```
Usage: psoracle_close oracle_connection
Disconnect from oracle.
```

## psoracle_createsql - Create SQL object with bind parameter
```
Usage: psoracle_createsql oracle_connection sql_command
Create SQL object with bind parameter.
```

## psoracle_bindsql - Bind real value to bind variable
```
Usage: psoracle_bindsql oracle_command name value
Bind real value to bind variable.
```

## psoracle_unbindsql - Unbind bind variable
```
Usage: psoracle_unbindsql oracle_command bind_object
Unbind bind variable.
```

## psoracle_execupdatesql - Exuceute SQL with update
```
Usage: psoracle_execupdatesql oracle_command
Exuceute SQL with update.
```

## psoracle_execsql - Exuceute SQL without update
```
Usage: psoracle_execsql oracle_command
Exuceute SQL without update.
```

## psoracle_fetch - Fetch row
```
Usage: psoracle_fetch ([Ref]oracle_reader)
Fetch row.
```

## psoracle_free - Free object for oracle access
```
Usage: psoracle_free oracle_object
Free object for oracle access.
```

## psoracle_begin - Begin transaction
```
Usage: psoracle_begin oracle_connection
Begin transaction.
```

## psoracle_settran - Set transaction for oracle command
```
Usage: psoracle_settran oracle_command oracle_transaction
Set transaction for oracle command.
```

## psoracle_rollback - rollback transaction
```
Usage: psoracle_rollback oracle_transaction
Rollback transaction.
```

## psoracle_commit - commit transaction
```
Usage: psoracle_commit oracle_transaction
Commit transaction.
```

## pssock_open - Open socket for client
```
Usage: pssock_open server_ip server_port [encoding]
Open socket for client.
```

## pssock_close - Close socket for client
```
Usage: pssock_close socket_param
Close socket for client.
```

## pssock_readline - Read a line from socket
```
Usage: pssock_readline socket_param
Read a line from socket.
```

## pssock_writeline - Write a line to socket
```
Usage: pssock_writeline socket_param string
Write a line to socket.
```

## pssock_read - Read data as binary from socket
```
Usage: pssock_read socket_param array_byte length
Read data as binary from socket.
```

## pssock_write - Write data as binary to socket
```
Usage: pssock_write socket_param array_byte length
Write data as binary to socket.
```

## pssock_start - Start server
```
Usage: pssock_start server_port
Start server.
```

## pssock_stop - Stop server
```
Usage: pssock_stop server
Stop server.
```

## pssock_accept - Accept connection from client
```
Usage: pssock_accept server [encoding]
Accept connection from client.
```

## pssock_accept - Unaccept(disconnect) connection from client
```
Usage: pssock_unaccept socket_param
Unaccept(disconnect) connection from client.
```

## psrunspc_getarraylist - Get System.Collections.ArrayList
```
Usage: psrunspc_getarraylist
Get System.Collections.ArrayList.
```

## psrunspc_open - Create and Open RunSpacePool
```
Usage: psrunspc_open max_runspace
Create and Open RunSpacePool.
```

## psrunspc_close - Close RunSpacePool
```
Usage: psrunspc_close run_space_pool
Close RunSpacePool.
```

## psrunspc_createthread - Create thread of powershell and add script to it
```
Usage: psrunspc_createthread run_space_pool script_block
Create thread of powershell and add script to it.
```

## psrunspc_addargument - Add argument to thread of powershell
```
Usage: psrunspc_addargument thread_of_powershellargument
Add argument to thread of powershell.
```

## psrunspc_begin - Begin script in thread
```
Usage: psrunspc_begin thread_of_powershell arraylist_of_powershell arraylist_of_child
Begin script in thread.
```

## psrunspc_wait - Wait terminate of all child thread
```
Usage: psrunspc_wait arraylist_of_powershell arraylist_of_child
Wait terminate of all child thread.
```

## psrunspc_waitasync - Wait asynchronously terminate of all child thread
```
Usage: psrunspc_waitasync arraylist_of_powershell arraylist_of_child
Wait asynchronously terminate of all child thread.
```
