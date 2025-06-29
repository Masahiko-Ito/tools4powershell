# tools4powershell

tools4powershell is a collection of some useful functions for powershell :)

## pscat - concatenate files and print on the standard output
```
Usage: pscat [-h|--help] [-n] [input ...]
Concatenate input(s), or standard input, to standard output.

  -e encoding    encoding for get-content called internally
  -n             number all output lines
```

## psgrep - print lines matching a pattern
```
Usage: psgrep [-h|--help] [-v] [-i] regex [input ...]
Search for regex in each input or standard input.

  -v             select non-matching lines
  -i             ignore case distinctions
  -e encoding    encoding for get-content called internally
```

## pswcl - print newline counts for each file
```
Usage: pswcl [-h|--help] [input ...]
Print newline counts for each input, and a total line if more than one input is specified.

  -e encoding        encoding for get-content called internally
```

## pssed - stream editor for filtering and transforming text
```
Usage: pssed [-h|--help] regex string [input ...]
For each substring matching regex in each lines from input, substitute the string.

  -e encoding        encoding for get-content called internally
```

## pshead - output the first part of files
```
Usage: pshead [-h|--help] [-l line_number] [input ...]
Print the first 10 lines of each input to standard output.
With no input, read standard input.

  -l line_number        print the first line_number lines instead of the first 10
  -e encoding           encoding for get-content called internally
```

## pstail - output the last part of files
```
Usage: pstail [-h|--help] [-l line_number] [input ...]
Print the last 10 lines of each input to standard output.
With no input, read standard input.

  -l line_number        print the last line_number lines instead of the last 10
  -e encoding           encoding for get-content called internally
```

## pscut - remove sections from each line of files
```
Usage: pscut [-h|--help] [-d "delimiter"] -i "index,..." [input ...]
Print selected parts of lines from each input to standard output.
With no input, read standard input.

  -d "delimiter"        use "delimiter" instead of "," for field delimiter
  -i "index,..."        select only these fields(0 origin)
  -e encoding           encoding for get-content called internally
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

  -d             only print duplicate lines
  -c             prefix lines by the number of occurrences
  -e encoding    encoding for get-content called internally
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

  -e encoding       encoding for get-content called internally

BUGS
  If input is not specified and output is specified, only last excel sheet in current directory will remain in output file.
```

## psxls2sheetname - get sheetname from excel book
```
Usage: psxls2sheetname [-h|--help] [-i input] [-s sheet]
Print sheet name identified by sheet.
If sheet is omitted, print all sheet names.
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
ex. Handle as text
  $inObj = psopen -r "input.txt"
  $outObj = psopen -w "output.txt"
  while (($rec = $inObj.readLine()) -ne $null){
      $outObj.writeLine($rec)
  }
  $inObj.close()
  $outObj.close()
  
ex. Handle data as binary
  $arrayByte = New-Object byte[] (1024)
  $inObj = psopen -r "input.txt"
  $outObj = psopen -w "output.txt"
  $read_length = $inObj.BaseStream.Read($arrayByte, 0, $arrayByte.length)
  while ($read_length -gt 0){
      $outObj.BaseStream.Write($arrayByte, 0, $read_length)
      $read_length = $inObj.BaseStream.Read($arrayByte, 0, $arrayByte.length)
  }
  $inObj.close()
  $outObj.close() 
```

## psconvenc - Convert encoding of text file
```
Usage: psconvenc -i inputfile -ie encoding -o outputfile -oe encoding
```

## psexcel_open - Open excel_file and get excel_object for it
```
Usage: psexcel_open excel_file
Open excel_file and get excel_object for it.
ex.
    $xls = psexcel_open "Foo.xlsx"
```

## psexcel_update - Save by overwrite
```
Usage: psexcel_update excel_object
Save by overwrite.
ex.
    psexcel_update $xls
```

## psexcel_save - Save into excel_file
```
Usage: psexcel_save excel_object excel_file
Save into excel_file.
ex.
    psexcel_save $xls "Foo.xlsx"
```

## psexcel_close - Close excel file by excel_object
```
Usage: psexcel_close excel_object
Close excel file by excel_object.
ex.
    psexcel_close $xls
```

## psexcel_getCell - Get value from range on sheet
```
Usage: psexcel_getCell excel_object sheet range
Get value from range on sheet.
ex.
    $val1 = psexcel_getCell $xls "Sheet1" "A1"
    $val2 = psexcel_getCell $xls 2 "B2"
```

## psexcel_setCell Set value on range on sheet
```
Usage: psexcel_setCell excel_object sheet range value
Set value on range on sheet.
ex.
    psexcel_setCell $xls "Sheet1" "A1" "some text"
    psexcel_setCell $xls 2 "C3" "=A1+B2"
```

## psexcel_getFormula - Get formula from range on sheet
```
Usage: psexcel_getFormula excel_object sheet range
Get formula from range on sheet.
ex.
    $f1 = psexcel_getFormula $xls "Sheet1" "A1"
    $f2 = psexcel_getFormula $xls 2 "B2"
```

## psexcel_setFormula - Set formula on range on sheet
```
Usage: psexcel_setFormula excel_object sheet range formula
Set formula on range on sheet.
ex.
    psexcel_setFormula $xls "Sheet1" "A1" "=A1+B2"
    psexcel_setFormula $xls 2 "B2" "=C3+D4"
```

## psexcel_getBackgroundColor - Get index of background color from range on sheet
```
Usage: psexcel_getBackgroundColor excel_object sheet range
Get index of background color from range on sheet.
ex.
    $colIndex1 = psexcel_getBackgroundColor $xls "Sheet1" "A1"
    $colIndex2 = psexcel_getBackgroundColor $xls 2 "B2"
```

## psexcel_setBackgroundColor - Set background color to color_index on range on sheet
```
Usage: psexcel_setBackgroundColor excel_object sheet range color_index
Set background color to color_index on range on sheet.
ex.
    psexcel_setBackgroundColor $xls "Sheet1" "A1" 1
    psexcel_setBackgroundColor $xls 2 "B2" 3
```

## psexcel_getForegroundColor Get index of foreground color from range on sheet
```
Usage: psexcel_getForegroundColor excel_object sheet range
Get index of foreground color from range on sheet.
ex.
    $colIndex1 = psexcel_getForegroundColor $xls "Sheet1" "A1"
    $colIndex2 = psexcel_getForegroundColor $xls 2 "B2"
```

## psexcel_setForegroundColor - Set foreground color to color_index on range on sheet
```
Usage: psexcel_setForegroundColor excel_object sheet range color_index
Set foreground color to color_index on range on sheet.
ex.
    psexcel_setForegroundColor $xls "Sheet1" "A1" 1
    psexcel_setForegroundColor $xls 2 "B2" 3
```

## psexcel_getBold - Get $true or $false about bold from range on sheet
```
Usage: psexcel_getBold excel_object sheet range
Get $true or $false about bold from range on sheet.
ex.
    $boolean1 = psexcel_getBold $xls "Sheet1" "A1"
    $boolean2 = psexcel_getBold $xls 2 "B2"
```

## psexcel_turnonBold - Turn on bold on range on sheet
```
Usage: psexcel_turnonBold excel_object sheet range
Turn on bold on range on sheet.
ex.
    psexcel_turnonBold $xls "Sheet1" "A1"
    psexcel_turnonBold $xls 2 "B2"
```

## psexcel_turnoffBold - Turn off bold on range on sheet
```
Usage: psexcel_turnoffBold excel_object sheet range
Turn off bold on range on sheet.
ex.
    psexcel_turnoffBold $xls "Sheet1" "A1
    psexcel_turnoffBold $xls 2 "B2"
```

## psexcel_toggleBold - Toggle bold on range on sheet
```
Usage: psexcel_toggleBold excel_object sheet range
Toggle bold on range on sheet.
ex.
    psexcel_toggleBold $xls "Sheet1" "A1"
    psexcel_toggleBold $xls 2 "B2"
```

## psexcel_getSheetCount - Get count of sheets
```
Usage: psexcel_getSheetCount excel_object
Get count of sheets.
ex.
    $sc = psexcel_getSheetCount $xls
```

## psexcel_getSheetName - Get name of sheet
```
Usage: psexcel_getSheetName excel_object sheet
Get name of sheet.
ex.
    $sn = psexcel_getSheetName $xls 2
```

## psexcel_setSheetName - Set name of sheet
```
Usage: psexcel_setSheetName excel_object sheet name
Set name of sheet.
ex.
    psexcel_setSheetName $xls "Sheet1" "SHEET1"
    psexcel_setSheetName $xls 2 "SHEET2"
```

## psexcel_getActiveSheetName Get name of active sheet
```
Usage: psexcel_getActiveSheetName excel_object
Get name of active sheet.
ex.
    $sn = psexcel_getActiveSheetName $xls
```

## psexcel_setActiveSheetName - Set name of active sheet
```
Usage: psexcel_setActiveSheetName excel_object name
Set name of active sheet.
ex.
     psexcel_setActiveSheetName $xls "SHEET1"
```

## psexcel_copyCell - Copy range of cell to another cell
```
Usage: psexcel_copyCell excel_object source_sheet source_range dest_sheet dest_cell
Copy range of cell to another cell.
ex.
    psexcel_copyCell $xls "Sheet1" "A1:C3" "sheet2" "D4"
    psexcel_copyCell $xls 1 "A1:C3" 2 "D4"
```

## psexcel_replaceString - Replace string to another string
```
Usage: psexcel_replaceString excel_object sheet range before_string after_string
Replace string to another string.
ex.
    psexcel_replaceString $xls "Sheet1" "A1:C3" "#1#" "Hoge"
    psexcel_replaceString $xls 1 "A1:C3" "#1#" "Hoge"
```

## psexcel_preview - Print preview sheet
```
Usage: psexcel_preview excel_object sheet
Print preview sheet.
ex.
    psexcel_preview $xls "Sheet1"
    psexcel_preview $xls 2
```

## psexcel_print - Print sheet
```
Usage: psexcel_print excel_object sheet
Print sheet.
ex.
    psexcel_print $xls "Sheet1"
    psexcel_print $xls 2
```

## psexcel_turnonVisible - Turn on visible
```
Usage: psexcel_turnonVisible excel_object
Turn on visible.
ex.
    psexcel_turnonVisible $xls
```

## psexcel_turnoffVisible - Turn off visible
```
Usage: psexcel_turnoffVisible excel_object
Turn off visible.
ex.
    psexcel_turnoffVisible $xls
```

## psexcel_turnonAlert - Turn on displayAlerts
```
Usage: psexcel_turnonAlert excel_object
Turn on displayAlerts.
ex.
    psexcel_turnonAlert $xls
```

## psexcel_turnoffAlert - Turn off displayAlerts
```
Usage: psexcel_turnoffAlert excel_object
Turn off displayAlerts.
ex.
    psexcel_turnoffAlert $xls
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
ex.
    $ocon = psoracle_open "127.0.0.1" "taro" "himitsu"
```

## psoracle_close - Disconnect from oracle
```
Usage: psoracle_close oracle_connection
Disconnect from oracle.
ex.
    psoracle_close $ocon
```

## psoracle_createsql - Create SQL object with bind parameter
```
Usage: psoracle_createsql oracle_connection sql_command
Create SQL object with bind parameter.
ex.
    $ocmd = psoracle_createsql $ocon "select * from table where id=:id_value"
    ... something to do ...
    psoracle_free $ocmd
```

## psoracle_bindsql - Bind real value to bind variable
```
Usage: psoracle_bindsql oracle_command name value
Bind real value to bind variable.
ex.
    $id_value = psoracle_bindsql $ocmd "id_value" "12345"
```

## psoracle_unbindsql - Unbind bind variable
```
Usage: psoracle_unbindsql oracle_command bind_object
Unbind bind variable.
ex.
    psoracle_unbindsql $ocmd $id_value
```

## psoracle_execupdatesql - Exuceute SQL with update
```
Usage: psoracle_execupdatesql oracle_command
Exuceute SQL with update.
ex.
    $count = psoracle_execupdatesql $ocmd
    if ($count -lt 0){
        write-output "Error: update SQL"
    }
```

## psoracle_execsql - Exuceute SQL without update
```
Usage: psoracle_execsql oracle_command
Exuceute SQL without update.
ex.
    $ocr = psoracle_execsql $ocmd
    ... something to do ...
    psoracle_free $ocr
```

## psoracle_fetch - Fetch row
```
Usage: psoracle_fetch ([Ref]oracle_reader)
Fetch row.
ex.
    while (psoracle_fetch ([Ref]$ocr)){
        write-output $ocr["id"]
    }
```

## psoracle_free - Free object for oracle access
```
Usage: psoracle_free oracle_object
Free object for oracle access.
ex.
    psoracle_free $ocr
    psoracle_free $ocmd
```

## psoracle_begin - Begin transaction
```
Usage: psoracle_begin oracle_connection
Begin transaction.
ex.
    $otran = psoracle_begin $ocon
```

## psoracle_settran - Set transaction for oracle command
```
Usage: psoracle_settran oracle_command oracle_transaction
Set transaction for oracle command.
ex.
    psoracle_settran $ocmd $otran
```

## psoracle_rollback - rollback transaction
```
Usage: psoracle_rollback oracle_transaction
Rollback transaction.
ex.
    psoracle_rollback $otran
```

## psoracle_commit - commit transaction
```
Usage: psoracle_commit oracle_transaction
Commit transaction.
ex.
    psoracle_commit $otran
```

## pssock_open - Open socket for client
```
Usage: pssock_open server_ip server_port [encoding]
Open socket for client.
ex.
    $param = pssock_open "127.0.0.1" "12345" "UTF-8"
```

## pssock_close - Close socket for client
```
Usage: pssock_close socket_param
Close socket for client.
ex.
    pssock_close $param
```

## pssock_readline - Read a line from socket
```
Usage: pssock_readline socket_param
Read a line from socket.
ex.
    $line = pssock_readline $param
```

## pssock_writeline - Write a line to socket
```
Usage: pssock_writeline socket_param string
Write a line to socket.
ex.
    $stat = pssock_writeline $param $line
```

## pssock_read - Read data as binary from socket
```
Usage: pssock_read socket_param array_byte length
Read data as binary from socket.
ex.
    $arrayByte = New-Object byte[] (1024)
    $read_length = pssock_read $param $arrayByte $arrayByte.length
```

## pssock_write - Write data as binary to socket
```
Usage: pssock_write socket_param array_byte length
Write data as binary to socket.
ex.
    $arrayByte = New-Object byte[] (1024)
    $stat = pssock_write $param $arrayByte $arrayByte.length
```

## pssock_start - Start server
```
Usage: pssock_start server_port
Start server.
ex.
    $server = pssock_start "12345"
```

## pssock_stop - Stop server
```
Usage: pssock_stop server
Stop server.
ex.
    pssock_stop $server
```

## pssock_accept - Accept connection from client
```
Usage: pssock_accept server [encoding]
Accept connection from client.
ex.
    $param = pssock_accept $server "UTF-8"
```

## pssock_accept - Unaccept(disconnect) connection from client
```
Usage: pssock_unaccept socket_param
Unaccept(disconnect) connection from client.
ex.
    pssock_unaccept $param
```

## pssock_getip - Get ip-address of client
```
Usage: pssock_getip socket_param
Get ip-address of client.
ex.
    $ip = pssock_getip $param
```

## pssock_getipstr - Get ip-address string of client
```
Usage: pssock_getipstr socket_param
Get ip-address string of client.
ex.
    $ip = pssock_getipstr $param
```

## pssock_getport - Get port of client
```
Usage: pssock_getport socket_param
Get port of client.
ex.
    $ip = pssock_getport $param
```

## pssock_getportstr - Get port string of client
```
Usage: pssock_getportstr socket_param
Get port string of client.
ex.
    $ip = pssock_getportstr $param
```

## psrunspc_getarraylist - Get System.Collections.ArrayList
```
Usage: psrunspc_getarraylist
Get System.Collections.ArrayList.
ex.
    $refArray = psrunspc_getarraylist
    $Array0 = $ref.Array.value[0]
```

## psrunspc_open - Create and Open RunSpacePool
```
Usage: psrunspc_open max_runspace
Create and Open RunSpacePool.
ex.
    $rsp = psrunspc_open 10
    ... something to do ...
    psrunspc_close $rsp
```

## psrunspc_close - Close RunSpacePool
```
Usage: psrunspc_close run_space_pool
Close RunSpacePool.
ex.
    $rsp = psrunspc_open 10
    ... something to do ...
    psrunspc_close $rsp
```

## psrunspc_createthread - Create thread of powershell and add script to it
```
Usage: psrunspc_createthread run_space_pool script_block
Create thread of powershell and add script to it.
ex.
    $ps = psrunspc_createthread $rsp $script
```

## psrunspc_addargument - Add argument to thread of powershell
```
Usage: psrunspc_addargument thread_of_powershellargument
Add argument to thread of powershell.
ex.
    psrunspc_addargument $argument
```

## psrunspc_begin - Begin script in thread
```
Usage: psrunspc_begin thread_of_powershell arraylist_of_powershell arraylist_of_child
Begin script in thread.
ex.
    $refAryps = psrunspc_getarraylist
    $refArych = psrunspc_getarraylist
    ... something to do ...
    psrunspc_begin $ps $refAryps.value $refArych.value
```

## psrunspc_wait - Wait terminate of all child thread
```
Usage: psrunspc_wait arraylist_of_powershell arraylist_of_child
Wait terminate of all child thread.
ex.
    psrunspc_wait $refAryps.value $refArych.value
```

## psrunspc_waitasync - Wait asynchronously terminate of all child thread
```
Usage: psrunspc_waitasync arraylist_of_powershell arraylist_of_child
Wait asynchronously terminate of all child thread.
ex.
    psrunspc_waitasync $refAryps.value $refArych.value
```

## psrpa_init - Initialize rpa environment
```
Usage: psrpa_init
Initialize rpa environment.
ex.
    $rpa = psrpa_init
```

## psrpa_setBeforeWait - Set a time(ms) for waiting before functions(psrpa_*)
```
Usage: psrpa_setBeforeWait rpa_object wait_time_ms
Set a time(ms) for waiting before functions(psrpa_*).
```

## psrpa_setAfterWait - Set a time(ms) for waiting after functions(psrpa_*)
```
Usage: psrpa_setAfterWait rpa_object wait_time_ms
Set a time(ms) for waiting after functions(psrpa_*).
```

## psrpa_showMouse - Show current mouse position for debug purpose
```
Usage: psrpa_showMouse rpa_object
Show current mouse position for debug purpose.
```

## psrpa_showMouseByClick - Show current mouse position for psrpa_setMouse and psrpa_clickPoint by click
```
Usage: psrpa_showMouseByClick rpa_object [wait_sec]
Show current mouse position for psrpa_setMouse and psrpa_clickPoint by click.
Press any key to terminate.
    wait_sec    default 5
ex.
    $wait_sec = 10
    psrpa_showMouseByClick $rpa $wait_sec
```

## psrpa_setMouse - Set mouse position
```
Usage: psrpa_setMouse rpa_object x_position y_position
Set mouse position.
ex.
    $x = 10
    $y = 20
    psrpa_setMouse $rpa $x $y
```

## psrpa_click - Click mouse button
```
Usage: psrpa_click rpa_object mouse_button click_action
Click mouse button.
    mouse_button left, LEFT, l, L
                 middle, MIDDLE, m, M
                 right, RIGHT, r, R
    click_action click, CLICK, 1
                 2click, 2CLICK, 2
                 3click, 3CLICK, 3
                 down, DOWN, d, D
                 up, UP, u, U
ex.
    psrpa_click $rpa "left" "click"
```

## psrpa_clickPoint - Set mouse position and click mouse button
```
Usage: psrpa_clickPoint rpa_object x_position y_position mouse_button click_action
Set mouse position and click mouse button.
    mouse_button left, LEFT, l, L
                 middle, MIDDLE, m, M
                 right, RIGHT, r, R
    click_action click, CLICK, 1
                 2click, 2CLICK, 2
                 3click, 3CLICK, 3
                 down, DOWN, d, D
                 up, UP, u, U
ex.
    $x = 10
    $y = 20
    psrpa_clickPoint $rpa $x $y "left" "click"
```

## psrpa_showAppTitle - Show application and title
```
Usage: psrpa_showAppTitle rpa_object
Show application and title.
```

## psrpa_activateWindow - Activate specified window
```
Usage: psrpa_activateWindow rpa_object application title
Activate specified window
    application    "" means "^$", $null means "^.*$"
    title          "" means "^$", $null means "^.*$"
ex.
    $application = "notepad"
    $title = "no title - memo pad"
    psrpa_activateWindow $rpa $application $title
```

## psrpa_setWindow - Set window position and size
```
Usage: psrpa_setWindow application rpa_object title x_position y_position width height
Set window position and size
    application    "" means "^$", $null means "^.*$"
    title          "" means "^$", $null means "^.*$"
ex.
    $application = "notepad"
    $title = "no title - memo pad"
    $x = 10
    $y = 20
    $width = 300
    $height = 400
    psrpa_setWindow $rpa $application $title $x $y $width $height
```

## psrpa_sendKeys - Send keys(string, function, special keys etc)
```
Usage: psrpa_sendKeys rpa_object string
Send keys(string, function, special keys etc)
    string any-string
           {BS} for backspace
           {DEL} for delete
           {PRTSC} for print screen
           {ENTER} for enter
           {LEFT} for left arrow
           {F1} for F1
           ^a for CTRL-a
           +a for SHIFT-a
           %a for ALT-a
           other keys - search by google :P
ex.
    psrpa_sendKeys $rpa "any string"
    psrpa_sendKeys $rpa "{BS}"
    psrpa_sendKeys $rpa "^a"
```

## psrpa_sendKeyEx - Send a key(Both of normal key and extended key are acceptable)
```
Usage: psrpa_sendKeyEx rpa_object virtual_keycode action isExtended
Send a key(Both of normal key and extended key are acceptable)
    virtual_keycode see https://msdn.microsoft.com/ja-jp/windows/desktop/dd375731
    action "down", "DOWN", "d", "D"
           "up", "UP", "u", "U"
           
    isExtended $true or $false
               see https://learn.microsoft.com/en-us/windows/win32/inputdev/about-keyboard-input#extended-key-flag
                   https://kusidate.blog.fc2.com/blog-entry-31.html
ex.
    $right_ctrl = 0xa3
    psrpa_sendKeyEx $rpa $right_ctrl "send" $true
```

## psrpa_setClipboard - Set clipboard to string
```
Usage: psrpa_setClipboard rpa_object string
Set clipboard to string.
Caution!
    powershell.exe have to be invoked with -sta option.
ex.
    psrpa_setClipboard $rpa "any string"
```

## psrpa_getClipboard - Get string from clipboard
```
Usage: psrpa_getClipboard rpa_object
Get string from clipboard.
Caution!
    powershell.exe have to be invoked with -sta option.
ex.
    $str = psrpa_getClipboard $rpa
```

## psrpa_getBmp - Get image into file(BMP)
```
Usage: psrpa_getBmp rpa_object left_x top_x right_x bottom_y output.bmp
Get image into file(BMP).
Output file name will be (left_x)_(top_y)_(right_x)_(bottom_y)_output.bmp.
ex.
    psrpa_getBmp $rpa 10 10 200 100 "icon.bmp"
```

## psrpa_getBmpByClick - Get image into file(BMP) by click
```
Usage: psrpa_getBmpByClick rpa_object output.bmp [wait_sec]
Get image into file(BMP) by click.
Output file name will be (left_x)_(top_y)_(right_x)_(bottom_y)_output.bmp.
Mouse cursor is CROSS for pointing (left_x, top_y).
                HAND  for pointing (right_x, bottom_y).
Press any key to terminate.
    wait_sec    default 5
ex.
    $wait_sec = 10
    psrpa_getBmpByClick $rpa "icon.bmp" $wait_sec
```

## psrpa_compareBmp - Compare specifoed rectangle and bmpfile
```
Usage: psrpa_compareBmp rpa_object left_x top_x right_x bottom_y input.bmp
Compare specifoed rectangle and bmpfile.
Return $true when specified rectangle and input.bmp are same.
ex.
    $bool = psrpa_compareBmp $rpa 10 10 200 100 "icon.bmp"
    if ($bool){
        write-output "match"
    }else{
        write-output "not match"
    }
```

## psrpa_isSameBmp - Compare rectangle starts with specified point and bmpfile
```
Usage: psrpa_isSameBmp rpa_object left top input.bmp
Compare rectangle starts with specified point and bmpfile.
Return $true when same bmp.
ex.
    $bool = psrpa_isSameBmp $rpa 10 10 "icon.bmp"
    if ($bool){
        write-output "match"
    }else{
        write-output "not match"
    }
```

## psrpa_searchBmp - Search bmpfile in screen
```
Usage: psrpa_searchBmp rpa_object left_x top_x right_x bottom_y input.bmp
Search bmpfile in screen.
Return @(left,top,right,bottom) when bmpfile is found in screen.
ex.
    $pos = psrpa_searchBmp $rpa $null $null $null $null "icon.bmp"
    if ($pos[0] -lt 0){
        write-output "not found"
    }else{
        write-output "found"
        $left = $pos[0]
        $top = $pos[1]
        $right = $pos[2]
        $bottom = $pos[3]
    }
```

## psrpa_searchBmpPosition - Get position of bmpfile in whole screen
```
Usage: psrpa_searchBmpPosition rpa_object input.bmp
Get position of bmpfile in whole screen.
Return @(left,top,width,height) when bmpfile is found in screen.
ex.
    $pos = psrpa_searchBmpPosition $rpa "icon.bmp"
    if ($pos[0] -lt 0){
        write-output "not found"
    }else{
        write-output "found"
        $left = $pos[0]
        $top = $pos[1]
        $width = $pos[2]
        $height = $pos[3]
    }
```

## psrpa_waitBmp - Wait appearnce of specified bmp
```
Usage: psrpa_waitBmp rpa_object bmp_file
Wait appearnce of specified bmp.

ex.
    psrpa_waitBmp $rpa "icon.bmp"
```

## psrpa_waitBmpVanished - Wait vanishing of specified bmp
```
Usage: psrpa_waitBmpVanished rpa_object bmp_file
Wait vanishing of specified bmp.

ex.
    psrpa_waitBmpVanished $rpa "icon.bmp"
```

## psrpa_clickBmp - Click specified bmp
```
Usage: psrpa_clickBmp rpa_object bmp_file button click
Click specified bmp.

ex.
    psrpa_clickBmp $rpa "icon.bmp" "left" "click"
```

## psrpa_uia_show_element - Show child ui-automation element information in specified elemment
```
Usage: psrpa_uia_show_element rpa_object element
Show child ui-automation element information in specified elemment.
    element        "" or $null mean root-window
ex.
    psrpa_uia_show_element $rpa $element
```

## psrpa_uia_show_element_recursive - Show recursively child ui-automation element information in specified elemment
```
Usage: psrpa_uia_show_element_recursive rpa_object element
Show recursively child ui-automation element information in specified elemment.
    element        "" or $null mean root-window
ex.
    psrpa_uia_show_element_recursive $rpa $element
```

## psrpa_uia_show_element_all - Show all ui-automation element information in specified elemment
```
Usage: psrpa_uia_show_element_all rpa_object element
Show all ui-automation element information in specified elemment.
    element        "" or $null mean root-window
ex.
    psrpa_uia_show_element_all $rpa $element

This is obsolete. Use psrpa_uia_show_element_recursive instead.
```

## psrpa_uia_show - Show all ui-automation element information in root-window
```
Usage: psrpa_uia_show rpa_object
Show all ui-automation element information in root-window.
ex.
    psrpa_uia_show $rpa

This is obsolete. Use psrpa_uia_show_element_recursive instead.
```

## psrpa_uia_get_element - Get child ui-automation element in specfied element
```
Usage: psrpa_uia_get_element rpa_object parent_element class_name localized_controlname name
Get child ui-automation element.
    parent_element        "" or $null mean root-window
    class_name            "" means "^$", $null means "^.*$"
    localized_controlname
    name
ex.
    $app = psrpa_uia_get_element $rpa $null "Notepad" "ウィンドウ" "無題 - メモ帳"
    $help = psrpa_uia_get_element $rpa $app "" "メニュー項目" "ヘルプ\(H\)"
```

## psrpa_uia_get_element_all - Get all ui-automation element in specfied element
```
Usage: psrpa_uia_get_element_all rpa_object parent_element class_name localized_controlname name
Get all ui-automation element in specfied element.
    parent_element        "" or $null mean root-window
    class_name            "" means "^$", $null means "^.*$"
    localized_controlname
    name
ex.
    $app = psrpa_uia_get_element_all $rpa $null "Notepad" "ウィンドウ" "無題 - メモ帳"
    $help = psrpa_uia_get_element_all $rpa $app "" "メニュー項目" "ヘルプ\(H\)"
```

## psrpa_uia_get - Get all ui-automation element in root window
```
Usage: psrpa_uia_get rpa_object parent_element class_name localized_controlname name
Get all ui-automation element in root window.
    class_name            "" means "^$", $null means "^.*$"
    localized_controlname
    name
ex.
    $app = psrpa_uia_get $rpa "Notepad" "ウィンドウ" "無題 - メモ帳"
    $help = psrpa_uia_get $rpa "" "メニュー項目" "ヘルプ\(H\)"
```

## psrpa_uia_getGeometry - Get geometry of ui-automation element(*Pattern)
```
Usage: psrpa_uia_getGeometry rpa_object element
Get geometry of ui-automation element(*Pattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $geo = psrpa_uia_getGeometry $rpa $elm
    $x = $geo[0]
    $y = $geo[1]
    $width = $geo[2]
    $height = $geo[3]
```

## psrpa_uia_invoke - Invoke ui-automation element(InvokePattern)
```
Usage: psrpa_uia_invoke rpa_object element
Invoke ui-automation element(InvokePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_invoke $rpa $elm
```

## psrpa_uia_expand - Expand ui-automation element(ExpandCollapsePattern)
```
Usage: psrpa_uia_expand rpa_object element
Expand ui-automation element(ExpandCollapsePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_expand $rpa $elm
```

## psrpa_uia_close - Close ui-automation element(WindowPattern)
```
Usage: psrpa_uia_close rpa_object element
Close ui-automation element(WindowPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_close $rpa $elm
```

## psrpa_uia_setValue - Set value into ui-automation element(ValuePattern)
```
Usage: psrpa_uia_setValue rpa_object element value
Set value into ui-automation element(ValuePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_setValue $rpa $elm "any-string"
```

## psrpa_uia_getText - Get text from ui-automation element(TextPattern)
```
Usage: psrpa_uia_getText rpa_object element
Get text from ui-automation element(TextPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $text = psrpa_uia_getText $rpa $elm
```

## psrpa_uia_toggle - Toggle ui-automation element(TogglePattern)
```
Usage: psrpa_uia_toggle rpa_object element
Toggle ui-automation element(TogglePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_toggle $rpa $elm
```

## psrpa_uia_select - Select ui-automation element(SelectionItem)
```
Usage: psrpa_uia_select rpa_object element
Select ui-automation element(SelectionItem).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_select $rpa $elm
```

## psrpa_uia_move - Move ui-automation element(TransformPattern)
```
Usage: psrpa_uia_move rpa_object element x y
Move ui-automation element(TransformPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $x = 10
    $y = 20
    psrpa_uia_move $rpa $elm $x $y
```

## psrpa_uia_resize - Resize ui-automation element(TransformPattern)
```
Usage: psrpa_uia_resize rpa_object element width height
Resize ui-automation element(TransformPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $width = 100
    $height = 200
    psrpa_uia_resize $rpa $elm $width $height
```

## psrpa_uia_getGridClassName - Get grid classname of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getGridClassName rpa_object element row col
Get grid classname of ui-automation element(GridPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $row = 0
    $col = 0
    $classname = psrpa_uia_getGridClassName $rpa $elm $row $col
```

## psrpa_uia_getGridControlType - Get grid controltype of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getGridControlType rpa_object element row col
Get grid controltype of ui-automation element(GridPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $row = 0
    $col = 0
    $controltype = psrpa_uia_getGridControlType $rpa $elm $row $col
```

## psrpa_uia_getGridName - Get grid name of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getGridName rpa_object element row col
Get grid name of ui-automation element(GridPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $row = 0
    $col = 0
    $name = psrpa_uia_getGridName $rpa $elm $row $col
```

## psrpa_uia_getGridColumn - Get grid colmun of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getGridColumn rpa_object element
Get grid column of ui-automation element(GridPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $col = psrpa_uia_getGridColumn $rpa $elm
```

## psrpa_uia_getGridRow - Get grid row of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getGridRow rpa_object element
Get grid row of ui-automation element(GridPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $row = psrpa_uia_getGridRow $rpa $elm
```

## psrpa_uia_getGridItemColumn - Get griditem column of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getGridItemColumn rpa_object element
Get griditem column of ui-automation element(GridItemPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $column = psrpa_uia_getGridItemColumn $rpa $elm
```

## psrpa_uia_getGridItemColumnSpan - Get griditem column span of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getGridItemColumnSpan rpa_object element
Get griditem column span of ui-automation element(GridItemPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $columnspan = psrpa_uia_getGridItemColumnSpan $rpa $elm
```

## psrpa_uia_getGridItemRow - Get griditem row of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getGridItemRow rpa_object element
Get griditem row of ui-automation element(GridItemPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $row = psrpa_uia_getGridItemRow $rpa $elm
```

## psrpa_uia_getGridItemRowSpan - Get griditem row span of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getGridItemRowSpan rpa_object element
Get griditem row span of ui-automation element(GridItemPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $rowspan = psrpa_uia_getGridItemRowSpan $rpa $elm
```

## psrpa_uia_setRangeValue - Set RangeValue of ui-automation element(RangeValuePattern)
```
Usage: psrpa_uia_setRangeValue rpa_object element value
Set RangeValue of ui-automation element(RangeValuePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_setRangeValue $rpa $elm 10
```

## psrpa_uia_getRangeValue - Get RangeValue of ui-automation element(RangeValuePattern)
```
Usage: psrpa_uia_getRangeValue rpa_object element
Get RangeValue of ui-automation element(RangeValuePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $value = psrpa_uia_getRangeValue $rpa $elm
```

## psrpa_uia_getRangeValueMax - Get Maximum RangeValue of ui-automation element(RangeValuePattern)
```
Usage: psrpa_uia_getRangeValueMax rpa_object element
Get Maximum RangeValue of ui-automation element(RangeValuePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $max = psrpa_uia_getRangeValueMax $rpa $elm
```

## psrpa_uia_getRangeValueMin - Get Minimum RangeValue of ui-automation element(RangeValuePattern)
```
Usage: psrpa_uia_getRangeValueMin rpa_object element
Get Minimum RangeValue of ui-automation element(RangeValuePattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $min = psrpa_uia_getRangeValueMin $rpa $elm
```

## psrpa_uia_setScroll - Set Scroll of ui-automation element(ScrollPattern)
```
Usage: psrpa_uia_setScroll rpa_object element horizontal_percent vertical_percent
Set Scroll of ui-automation element(ScrollPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_setScroll $rpa $elm 10 100
```

## psrpa_uia_getScroll - Get Scroll of ui-automation element(ScrollPattern)
```
Usage: psrpa_uia_getScroll rpa_object element
Get Scroll of ui-automation element(ScrollPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $scroll = psrpa_uia_getScroll $rpa $elm
    $scroll_horizontal = $scroll[0]
    $scroll_vertical = $scroll[1]
```

## psrpa_uia_getScrollHview - Get Horizontal view size(%) of Scroll of ui-automation element(ScrollPattern)
```
Usage: psrpa_uia_getScrollHview rpa_object element
Get Horizontal view size(%) of Scroll of ui-automation element(ScrollPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $horizontal_view_percent = psrpa_uia_getScrollHview $rpa $elm
```

## psrpa_uia_getScrollVview - Get Vertical view size(%) of Scroll of ui-automation element(ScrollPattern)
```
Usage: psrpa_uia_getScrollVview rpa_object element
Get Vertical view size(%) of Scroll of ui-automation element(ScrollPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    $vertical_view_percent = psrpa_uia_getScrollVview $rpa $elm
```

## psrpa_uia_transformMove - Move ui-automation element(TransformPattern)
```
Usage: psrpa_uia_transformMove rpa_object element x y
Move ui-automation element(TransformPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_transformMove $rpa $elm 10 20
```

## psrpa_uia_transformResize - Resize ui-automation element(TransformPattern)
```
Usage: psrpa_uia_transformResize rpa_object element width height
Resize ui-automation element(TransformPattern).
ex.
    $elm = psrpa_uia_get ...snip...
    psrpa_uia_transformResize $rpa $elm 100 200
```

## psrpa_uia_getPattern - Get pattern from ui-automation element
```
Usage: psrpa_uia_getPattern rpa_object element pattern
Get pattern from ui-automation element.
    pattern "Invoke"
            "ExpandCollapse"
            "Window"
            "Value"
            "Text"
            "Toggle"
            "SelectionItem"
            "Grid"
            "GridItem"
            "Selection"
            "RangeValue"
            "Scroll"
            "Transform"
            "VirtualizedItem"
            "ItemContainer"
            "Selection"
            "MultipleView"
            "Table"
            "TableItem"
            "ScrollItem"
ex.
    $elm = psrpa_uia_get ...snip...
    $pattern = psrpa_uia_getPattern $rpa $elm "Invoke"
    $pattern.Invoke()
```

## Youtube
[PowerShell用便利機能集](https://youtu.be/6Evv53xYsNc)
