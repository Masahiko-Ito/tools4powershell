$ErrorActionPreference = 'Stop'

#
# Sample program to access PostgreSQL by mytoolkit.ps1
#

#
# nupkgから展開したDLLをロード
#
#  npgsql.4.0.13.nupkg
#  system.memory.4.5.3.nupkg
#  system.runtime.compilerservices.unsafe.4.5.3.nupkg
#  system.threading.tasks.extensions.4.5.2.nupkg
#
[void][reflection.assembly]::LoadFrom("${PSSCriptRoot}\Npgsql.dll")

# データベースに接続
$IPADDRESS = "127.0.0.1"
$USER = "mito"
$PASSWORD = "postgresql"
$DBNAME = "testdb"
$PORT = 5432
$con = pspgsql_open $IPADDRESS $USER $PASSWORD $DBNAME $PORT

#
# テーブルを削除
#
$com = pspgsql_createsql $con "drop table tb_sample"
try{
	pspgsql_execddl $com
}catch{
	write-output $_.Exception.Message
}

#
# テーブルを再作成
#
$com = pspgsql_createsql $con @"
	create table tb_sample(
		id	char(4)	primary key,
		name	varchar(20),
		age	decimal
	)
"@
pspgsql_execddl $com

#
# まとめてSQL文作成(複数箇所で使用あり)
#
$com_delete = pspgsql_createsql $con "delete from tb_sample where id >= :id"
$com_insert = pspgsql_createsql $con "insert into tb_sample (id, name, age) values(:id, :name, :age)"
$com_select = pspgsql_createsql $con "select id, name, age from tb_sample where id >= :id"
$com_update = pspgsql_createsql $con "update tb_sample set name = :name where id = :id"

#
# テーブルにデータ追加１
#
$transaction = pspgsql_begin $con
#pspgsql_settran $com_insert $transaction
pspgsql_clearbindsql $com_insert
pspgsql_bindsql $com_insert ":id" "0001"
pspgsql_bindsql $com_insert ":name" "伊藤　太郎"
pspgsql_bindsql $com_insert ":age" 20
try{
	$rowcount = pspgsql_execupdatesql $com_insert
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# テーブルにデータ追加２
#
$transaction = pspgsql_begin $con
#pspgsql_settran $com_insert $transaction
pspgsql_clearbindsql $com_insert
pspgsql_bindsql $com_insert ":id" "0002"
pspgsql_bindsql $com_insert ":name" "伊藤　次郎"
pspgsql_bindsql $com_insert ":age" 19
try{
	$rowcount = pspgsql_execupdatesql $com_insert
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# テーブルにデータ追加３
#
$transaction = pspgsql_begin $con
#pspgsql_settran $com_insert $transaction
pspgsql_clearbindsql $com_insert
pspgsql_bindsql $com_insert ":id" "0003"
pspgsql_bindsql $com_insert ":name" "伊藤　三郎"
pspgsql_bindsql $com_insert ":age" 18
try{
	$rowcount = pspgsql_execupdatesql $com_insert
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# わざと重複キーでインサートして、どんなエラーメッセージが出るか観察してみる...
#
$transaction = pspgsql_begin $con
#pspgsql_settran $com_insert $transaction
pspgsql_clearbindsql $com_insert
pspgsql_bindsql $com_insert ":id" "0003"
pspgsql_bindsql $com_insert ":name" "伊藤　三郎"
pspgsql_bindsql $com_insert ":age" 18
try{
	$rowcount = pspgsql_execupdatesql $com_insert
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
#	"0" 個の引数を指定して "ExecuteNonQuery" を呼び出し中に例外が発生しました: "制約 'PK__tb_sampl__3213E83F3E6BC8CA' の PRIMARY KEY 違反。オブジェクト 'dbo.tb_sample' には重複するキーを挿入できません。重複するキーの値は (00000007) です。
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# テーブルを参照
#
echo "== 追加後 =="
$transaction = pspgsql_begin $con
#pspgsql_settran $com_select $transaction
pspgsql_clearbindsql $com_select
pspgsql_bindsql $com_select ":id" "0000"
$reader = pspgsql_execsql $com_select
try{
	while (pspgsql_fetch $reader){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	pspgsql_free $reader
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_free $reader
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# テーブルを更新
#
$transaction = pspgsql_begin $con
#pspgsql_settran $com_update $transaction
pspgsql_clearbindsql $com_update
pspgsql_bindsql $com_update ":id" "0003"
$bind_name = pspgsql_bindsql $com_update ":name" "伊藤　花子"
try{
	$rowcount = pspgsql_execupdatesql $com_update
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# テーブルを参照
#
echo "== 更新後 =="
$transaction = pspgsql_begin $con
#pspgsql_settran $com_select $transaction
pspgsql_clearbindsql $com_select
pspgsql_bindsql $com_select ":id" "0000"
$reader = pspgsql_execsql $com_select
try{
	while (pspgsql_fetch $reader){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	pspgsql_free $reader
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_free $reader
	pspgsql_rollback $transaction
}
pspgsql_free $reader
pspgsql_free $transaction

#
# テーブルからデータ削除
#
$transaction = pspgsql_begin $con
#pspgsql_settran $com_delete $transaction
pspgsql_clearbindsql $com_delete
pspgsql_bindsql $com_delete ":id" "0000"
try{
	$rowcount = pspgsql_execupdatesql $com_delete
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_rollback $transaction
}
pspgsql_free $transaction

#
# テーブルを参照
#
echo "== 全件削除後 =="
$transaction = pspgsql_begin $con
#pspgsql_settran $com_select $transaction
pspgsql_clearbindsql $com_select
pspgsql_bindsql $com_select ":id" "0000"
$reader = pspgsql_execsql $com_select
try{
	while (pspgsql_fetch $reader){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	pspgsql_free $reader
	pspgsql_commit $transaction
}catch{
	write-output $_.Exception.Message
	pspgsql_free $reader
	pspgsql_rollback $transaction
}
pspgsql_free $reader
pspgsql_free $transaction

#
# データベースから切断
#		
pspgsql_close $con
