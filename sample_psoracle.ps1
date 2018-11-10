# Connect oracle
$oracleConnection = psoracle_open "IP_ADDRESS"  "USERID" "PASSWORD"

# Create SQL statement with bind function
$oracleCommand_droptable = psoracle_createsql $oracleConnection "drop table tb_sample"
$oracleCommand_createtable = psoracle_createsql $oracleConnection "create table tb_sample (id char(4), name nvarchar2(20))"
$oracleCommand_insert = psoracle_createsql $oracleConnection "insert into tb_sample names(id, name) values(:id_value, :name_value)"
$oracleCommand_select = psoracle_createsql $oracleConnection "select * from tb_sample where id like :id_value"

#============================================================
# Start transaction
$oracleTransaction = psoracle_begin $oracleConnection

# Execute SQL for update
$oracleCount = psoracle_execupdatesql $oracleCommand_droptable
if ($oracleCount -lt 0){
	write-output "Error: drop table"
	# Rollback
	psoracle_rollback $oracleTransaction
}else{
	# Commit
	psoracle_commit $oracleTransaction
}

#============================================================
# Start transaction
$oracleTransaction = psoracle_begin $oracleConnection

# Execute SQL for update
$oracleCount = psoracle_execupdatesql $oracleCommand_createtable
if ($oracleCount -lt 0){
	write-output "Error: create table"
	# Rollback
	psoracle_rollback $oracleTransaction
}else{
	# Commit
	psoracle_commit $oracleTransaction
}

#============================================================
# Start transaction
$oracleTransaction = psoracle_begin $oracleConnection

# Bind real value
$id_value = psoracle_bindsql $oracleCommand_insert "id_value" "0001"
$name_value = psoracle_bindsql $oracleCommand_insert "name_value" "Oracle Ichiro"

# Execute SQL for update
$oracleCount1 = psoracle_execupdatesql $oracleCommand_insert
if ($oracleCount1 -lt 0){
	write-output "Error: insert"
}

# Free bind
psoracle_unbindsql $oracleCommand_insert $id_value
psoracle_unbindsql $oracleCommand_insert $name_value

#------------------------------------------------------------
# Bind real value
$id_value = psoracle_bindsql $oracleCommand_insert "id_value" "0002"
$name_value = psoracle_bindsql $oracleCommand_insert "name_value" "Oracle Jiro"

# Execute SQL for update
$oracleCount2 = psoracle_execupdatesql $oracleCommand_insert
if ($oracleCount2 -lt 0){
	write-output "Error: insert"
}

# Free bind
psoracle_unbindsql $oracleCommand_insert $id_value
psoracle_unbindsql $oracleCommand_insert $name_value

#------------------------------------------------------------
# Bind real value
$id_value = psoracle_bindsql $oracleCommand_insert "id_value" "0003"
$name_value = psoracle_bindsql $oracleCommand_insert "name_value" "Oracle Hanako"

# Execute SQL for update
$oracleCount3 = psoracle_execupdatesql $oracleCommand_insert
if ($oracleCount3 -lt 0){
	write-output "Error: insert"
}

# Free bind
psoracle_unbindsql $oracleCommand_insert $id_value
psoracle_unbindsql $oracleCommand_insert $name_value

if ($oracleCount1 -lt 0 -or $oracleCount2 -lt 0 -or $oracleCount3 -lt 0){
	# Rollback
	psoracle_rollback $oracleTransaction
}else{
	# Commit
	psoracle_commit $oracleTransaction
}

#============================================================
# Start transaction
$oracleTransaction = psoracle_begin $oracleConnection

# Bind read value
$id_value = psoracle_bindsql $oracleCommand_select "id_value" "000%"

# Execute SQL for query
$oracleReader = psoracle_execsql $oracleCommand_select

# Free bind
psoracle_unbindsql $oracleCommand_select $id_value

# Fetch row
while(psoracle_fetch ([Ref]$oracleReader)) {
	$oracleReader["id"] + "," + $oracleReader["name"]
}

# Free oracle reader object
psoracle_free $oracleReader

# Free sql object
psoracle_free $oracleCommand_droptable
psoracle_free $oracleCommand_createtable
psoracle_free $oracleCommand_insert
psoracle_free $oracleCommand_select

# Commit
psoracle_commit $oracleTransaction

#============================================================
# Disconnect oracle
psoracle_close $oracleConnection
