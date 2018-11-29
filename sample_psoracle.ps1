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
psoracle_settran $oracleCommand_droptable $oracleTransaction

try{
	# Execute SQL for update
	$oracleCount = psoracle_execupdatesql $oracleCommand_droptable
	# Commit
	psoracle_commit $oracleTransaction
}catch{
	write-output ("Error: drop table " + $Error[0])
	# Rollback
	psoracle_rollback $oracleTransaction
}

#============================================================
# Start transaction
$oracleTransaction = psoracle_begin $oracleConnection
psoracle_settran $oracleCommand_createtable $oracleTransaction

try{
	# Execute SQL for update
	$oracleCount = psoracle_execupdatesql $oracleCommand_createtable
	# Commit
	psoracle_commit $oracleTransaction
}catch{
	write-output ("Error: create table " + $Error[0])
	# Rollback
	psoracle_rollback $oracleTransaction
}

#============================================================
# Start transaction
$oracleTransaction = psoracle_begin $oracleConnection
# Set Transaction
psoracle_settran $oracleCommand_insert $oracleTransaction

# Bind real value
$id_value = psoracle_bindsql $oracleCommand_insert "id_value" "0001"
$name_value = psoracle_bindsql $oracleCommand_insert "name_value" "Oracle Ichiro"

try{
	# Execute SQL for update
	$oracleCount1 = psoracle_execupdatesql $oracleCommand_insert
}catch{
	write-output ("Error: insert " + $Error[0])
	$oracleCount1 = -1
}

# Free bind
psoracle_unbindsql $oracleCommand_insert $id_value
psoracle_unbindsql $oracleCommand_insert $name_value

#------------------------------------------------------------
# Bind real value
$id_value = psoracle_bindsql $oracleCommand_insert "id_value" "0002"
$name_value = psoracle_bindsql $oracleCommand_insert "name_value" "Oracle Jiro"

try{
	# Execute SQL for update
	$oracleCount2 = psoracle_execupdatesql $oracleCommand_insert
}catch{
	write-output ("Error: insert " + $Error[0])
	$oracleCount2 = -1
}

# Free bind
psoracle_unbindsql $oracleCommand_insert $id_value
psoracle_unbindsql $oracleCommand_insert $name_value

#------------------------------------------------------------
# Bind real value
$id_value = psoracle_bindsql $oracleCommand_insert "id_value" "0003"
$name_value = psoracle_bindsql $oracleCommand_insert "name_value" "Oracle Hanako"

try{
	# Execute SQL for update
	$oracleCount3 = psoracle_execupdatesql $oracleCommand_insert
}catch{
	write-output ("Error: insert " + $Error[0])
	$oracleCount3 = -1
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
psoracle_settran $oracleCommand_select $oracleTransaction

# Bind read value
$id_value = psoracle_bindsql $oracleCommand_select "id_value" "000%"

try{
	# Execute SQL for query
	$oracleReader = psoracle_execsql $oracleCommand_select
}catch{
	write-output ("Error: select " + $Error[0])
}

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
