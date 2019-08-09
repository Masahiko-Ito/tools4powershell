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

## pssed - stream editor for filtering and transforming text
```
Usage: pssed [-h|--help] regex string [input ...]
For each substring matching regex in each lines from input, substitute the string.
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

## psxls2sheetname - get sheetname from excel book
```
Usage: psxls2sheetname [-h|--help] [-i input] [-s sheet]
Print sheet name identified by sheet.
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

## pssock_getip - Get ip-address of client
```
Usage: pssock_getip socket_param
Get ip-address of client.
```

## pssock_getipstr - Get ip-address string of client
```
Usage: pssock_getipstr socket_param
Get ip-address string of client.
```

## pssock_getport - Get port of client
```
Usage: pssock_getport socket_param
Get port of client.
```

## pssock_getportstr - Get port string of client
```
Usage: pssock_getportstr socket_param
Get port string of client.
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

## psrpa_init - Initialize rpa environment
```
Usage: psrpa_init
Initialize rpa environment.
```

## psrpa_set_beforewait - Set a time(ms) for waiting before functions(psrpa_*)
```
Usage: psrpa_set_beforewait rpa_object wait_time_ms
Set a time(ms) for waiting before functions(psrpa_*).
```

## psrpa_set_afterwait - Set a time(ms) for waiting after functions(psrpa_*)
```
Usage: psrpa_set_afterwait rpa_object wait_time_ms
Set a time(ms) for waiting after functions(psrpa_*).
```

## psrpa_show_mouse_position - Show current mouse position for debug purpose
```
Usage: psrpa_show_mouse_position rpa_object
Show current mouse position for debug purpose.
```

## psrpa_show_mouse_position_byclick - Show current mouse position for psrpa_set_mouse and psrpa_position_click by click
```
Usage: psrpa_show_mouse_position_byclick rpa_object [wait_sec]
Show current mouse position for psrpa_set_mouse and psrpa_position_click by click.
```

## psrpa_set_mouse - Set mouse position
```
Usage: psrpa_set_mouse rpa_object x_position y_position
Set mouse position.
```

## psrpa_click - Click mouse button
```
Usage: psrpa_click rpa_object mouse_button click_action
Click mouse button.
```

## psrpa_position_click - Set mouse position and click mouse button
```
Usage: psrpa_position_click rpa_object x_position y_position mouse_button click_action
Set mouse position and click mouse button.
```

## psrpa_show_apptitle - Show application and title
```
Usage: psrpa_show_apptitle rpa_object
Show application and title.
```

## psrpa_activate_window - Activate specified window
```
Usage: psrpa_activate_window rpa_object application title
Activate specified window
```

## psrpa_set_window - Set window position and size
```
Usage: psrpa_set_window application rpa_object title x_position y_position width height
Set window position and size
```

## psrpa_sendkeys - Send keys(string, function, special keys etc)
```
Usage: psrpa_sendkeys rpa_object string
Send keys(string, function, special keys etc)
```

## psrpa_sendkeyEX - Send a key(Both of normal key and extended key are acceptable)
```
Usage: psrpa_sendkeyEX rpa_object virtual_keycode action isExtended
Send a key(Both of normal key and extended key are acceptable)
```

## psrpa_set_clipboard - Set clipboard to string
```
Usage: psrpa_set_clipboard rpa_object string
Set clipboard to string.
```

## psrpa_get_clipboard - Get string from clipboard
```
Usage: psrpa_get_clipboard rpa_object
Get string from clipboard.
```

## psrpa_get_bmp - Get image into file(BMP)
```
Usage: psrpa_get_bmp rpa_object left_x top_x right_x bottom_y output.bmp
Get image into file(BMP).
```

## psrpa_get_bmp_byclick - Get image into file(BMP) by click
```
Usage: psrpa_get_bmp_byclick rpa_object output.bmp [wait_sec]
Get image into file(BMP) by click.
```

## psrpa_compare_bmp - Compare specifoed rectangle and bmpfile
```
Usage: psrpa_compare_bmp rpa_object left_x top_x right_x bottom_y input.bmp
Compare specifoed rectangle and bmpfile.
```

## psrpa_search_bmp - Search bmpfile in screen
```
Usage: psrpa_search_bmp rpa_object left_x top_x right_x bottom_y input.bmp
```

## psrpa_uia_show - Show all ui-automation element information
```
Usage: psrpa_uia_show rpa_object
Show all ui-automation element.
```

## psrpa_uia_get - Get ui-automation element
```
Usage: psrpa_uia_get rpa_object parent_element class_name localized_controlname name
Get ui-automation element.
```

## psrpa_uia_invoke - Invoke ui-automation element(InvokePattern)
```
Usage: psrpa_uia_invoke rpa_object element
Invoke ui-automation element(InvokePattern).
```

## psrpa_uia_expand - Expand ui-automation element(ExpandCollapsePattern)
```
Usage: psrpa_uia_expand rpa_object element
Expand ui-automation element(ExpandCollapsePattern).
```

## psrpa_uia_close - Close ui-automation element(WindowPattern)
```
Usage: psrpa_uia_close rpa_object element
Close ui-automation element(WindowPattern).
```

## psrpa_uia_setvalue - Set value into ui-automation element(ValuePattern)
```
Usage: psrpa_uia_setvalue rpa_object element value
Set value into ui-automation element(ValuePattern).
```

## psrpa_uia_gettext - Get text from ui-automation element(TextPattern)
```
Usage: psrpa_uia_gettext rpa_object element
Get text from ui-automation element(TextPattern).
```

## psrpa_uia_toggle - Toggle ui-automation element(TogglePattern)
```
Usage: psrpa_uia_toggle rpa_object element
Toggle ui-automation element(TogglePattern).
```

## psrpa_uia_select - Select ui-automation element(SelectionItem)
```
Usage: psrpa_uia_select rpa_object element
Select ui-automation element(SelectionItem).
```

## psrpa_uia_move - Move ui-automation element(TransformPattern)
```
Usage: psrpa_uia_move rpa_object element x y
Move ui-automation element(TransformPattern).
```

## psrpa_uia_resize - Resize ui-automation element(TransformPattern)
```
Usage: psrpa_uia_resize rpa_object element width height
Resize ui-automation element(TransformPattern).
```

## psrpa_uia_getgrid_classname - Get grid classname of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getgrid_classname rpa_object element row col
Get grid classname of ui-automation element(GridPattern).
```

## psrpa_uia_getgrid_controltype - Get grid controltype of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getgrid_controltype rpa_object element row col
Get grid controltype of ui-automation element(GridPattern).
```

## psrpa_uia_getgrid_name - Get grid name of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getgrid_name rpa_object element row col
Get grid name of ui-automation element(GridPattern).
```

## psrpa_uia_getgrid_column - Get grid colmun of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getgrid_column rpa_object element
Get grid column of ui-automation element(GridPattern).
```

## psrpa_uia_getgrid_row - Get grid row of ui-automation element(GridPattern)
```
Usage: psrpa_uia_getgrid_row rpa_object element
Get grid row of ui-automation element(GridPattern).
```

## psrpa_uia_getgriditem_column - Get griditem column of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getgriditem_column rpa_object element
Get griditem column of ui-automation element(GridItemPattern).
```

## psrpa_uia_getgriditem_columnspan - Get griditem column span of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getgriditem_columnspan rpa_object element
Get griditem column span of ui-automation element(GridItemPattern).
```

## psrpa_uia_getgriditem_row - Get griditem row of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getgriditem_row rpa_object element
Get griditem row of ui-automation element(GridItemPattern).
```

## psrpa_uia_getgriditem_rowspan - Get griditem row span of ui-automation element(GridItemPattern)
```
Usage: psrpa_uia_getgriditem_rowspan rpa_object element
Get griditem row span of ui-automation element(GridItemPattern).
```

## psrpa_uia_getpattern - Get pattern from ui-automation element
```
Usage: psrpa_uia_getpattern rpa_object element pattern
Get pattern from ui-automation element.
```

