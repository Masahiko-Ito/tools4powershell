$ErrorActionPreference = 'Stop'

#
# Sample program to access MSSQLServer by mytoolkit.ps1
#

#
# データベースに接続
#
#$DATASOURCE = @"
#(
#	DESCRIPTION = (
#		ADDRESS_LIST = (
#			ADDRESS = (PROTOCOL = TCP)(HOST = 127.0.0.1)(PORT = 1521)
#		)
#	)
#	(
#		CONNECT_DATA = (SERVICE_NAME = XE)
#	)
#)
#"@
$DATASOURCE = "127.0.0.1"
$USER = "mito"
$PASSWORD = "oracle"
$con = psoracle_open $DATASOURCE $USER $PASSWORD

#
# テーブルを削除
#
$com = psoracle_createsql $con "drop table tb_sample"
try{
	psoracle_execddl $com
}catch{
	write-output $_.Exception.Message
}

#
# テーブルを再作成
#
$com = psoracle_createsql $con @"
	create table tb_sample(
		id	char(4)	primary key,
		name	nvarchar2(20),
		age	decimal
	)
"@
psoracle_execddl $com

#
# まとめてSQL文作成(複数箇所で使用あり)
#
$com_delete = psoracle_createsql $con "delete from tb_sample where id >= :id"
$com_insert = psoracle_createsql $con "insert into tb_sample (id, name, age) values(:id, :name, :age)"
$com_select = psoracle_createsql $con "select id, name, age from tb_sample where id >= :id"
$com_update = psoracle_createsql $con "update tb_sample set name = :name where id = :id"

#
# テーブルにデータ追加１
#
$transaction = psoracle_begin $con
psoracle_settran $com_insert $transaction
psoracle_clearbindsql $com_insert
psoracle_bindsql $com_insert ":id" "0001"
psoracle_bindsql $com_insert ":name" "伊藤　太郎"
psoracle_bindsql $com_insert ":age" 20
try{
	$rowcount = psoracle_execupdatesql $com_insert
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# テーブルにデータ追加２
#
$transaction = psoracle_begin $con
psoracle_settran $com_insert $transaction
psoracle_clearbindsql $com_insert
psoracle_bindsql $com_insert ":id" "0002"
psoracle_bindsql $com_insert ":name" "伊藤　次郎"
psoracle_bindsql $com_insert ":age" 19
try{
	$rowcount = psoracle_execupdatesql $com_insert
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# テーブルにデータ追加３
#
$transaction = psoracle_begin $con
psoracle_settran $com_insert $transaction
psoracle_clearbindsql $com_insert
psoracle_bindsql $com_insert ":id" "0003"
psoracle_bindsql $com_insert ":name" "伊藤　三郎"
psoracle_bindsql $com_insert ":age" 18
try{
	$rowcount = psoracle_execupdatesql $com_insert
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# わざと重複キーでインサートして、どんなエラーメッセージが出るか観察してみる...
#
$transaction = psoracle_begin $con
psoracle_settran $com_insert $transaction
psoracle_clearbindsql $com_insert
psoracle_bindsql $com_insert ":id" "0003"
psoracle_bindsql $com_insert ":name" "伊藤　三郎"
psoracle_bindsql $com_insert ":age" 18
try{
	$rowcount = psoracle_execupdatesql $com_insert
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
#	"0" 個の引数を指定して "ExecuteNonQuery" を呼び出し中に例外が発生しました: "制約 'PK__tb_sampl__3213E83F3E6BC8CA' の PRIMARY KEY 違反。オブジェクト 'dbo.tb_sample' には重複するキーを挿入できません。重複するキーの値は (00000007) です。
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# テーブルを参照
#
echo "== 追加後 =="
$transaction = psoracle_begin $con
psoracle_settran $com_select $transaction
psoracle_clearbindsql $com_select
psoracle_bindsql $com_select ":id" "0000"
$reader = psoracle_execsql $com_select
try{
	while (psoracle_fetch ([Ref]$reader)){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	psoracle_free $reader
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_free $reader
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# テーブルを更新
#
$transaction = psoracle_begin $con
psoracle_settran $com_update $transaction
psoracle_clearbindsql $com_update
psoracle_bindsql $com_update ":id" "0003"
$bind_name = psoracle_bindsql $com_update ":name" "伊藤　花子"
try{
	$rowcount = psoracle_execupdatesql $com_update
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# テーブルを参照
#
echo "== 更新後 =="
$transaction = psoracle_begin $con
psoracle_settran $com_select $transaction
psoracle_clearbindsql $com_select
psoracle_bindsql $com_select ":id" "0000"
$reader = psoracle_execsql $com_select
try{
	while (psoracle_fetch ([Ref]$reader)){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	psoracle_free $reader
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_free $reader
	psoracle_rollback $transaction
}
psoracle_free $reader
psoracle_free $transaction

#
# テーブルからデータ削除
#
$transaction = psoracle_begin $con
psoracle_settran $com_delete $transaction
psoracle_clearbindsql $com_delete
psoracle_bindsql $com_delete ":id" "0000"
try{
	$rowcount = psoracle_execupdatesql $com_delete
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_rollback $transaction
}
psoracle_free $transaction

#
# テーブルを参照
#
echo "== 全件削除後 =="
$transaction = psoracle_begin $con
psoracle_settran $com_select $transaction
psoracle_clearbindsql $com_select
psoracle_bindsql $com_select ":id" "0000"
$reader = psoracle_execsql $com_select
try{
	while (psoracle_fetch ([Ref]$reader)){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	psoracle_free $reader
	psoracle_commit $transaction
}catch{
	write-output $_.Exception.Message
	psoracle_free $reader
	psoracle_rollback $transaction
}
psoracle_free $reader
psoracle_free $transaction

#
# データベースから切断
#		
psoracle_close $con
