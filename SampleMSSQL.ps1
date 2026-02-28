$ErrorActionPreference = 'Stop'

#
# Sample program to access MSSQLServer by mytoolkit.ps1
#

#
# データベースに接続
#
#$DATASOURCE = "127.0.0.1,1433"
$DATASOURCE = "127.0.0.1"
$USER = "mito"
$PASSWORD = "sqlserver"
$DBNAME = "testdb"
$con = psmssql_open $DATASOURCE $USER $PASSWORD $DBNAME

#
# テーブルを削除
#
$com = psmssql_createsql $con "drop table tb_sample"
try{
	psmssql_execddl $com
}catch{
	write-output $_.Exception.Message
}

#
# テーブルを再作成
#
$com = psmssql_createsql $con @"
	create table tb_sample(
		id	char(4)	primary key,
		name	nvarchar(20),
		age	decimal
	)
"@
psmssql_execddl $com

#
# まとめてSQL文作成(複数箇所で使用あり)
#
$com_delete = psmssql_createsql $con "delete from tb_sample where id >= @id"
$com_insert = psmssql_createsql $con "insert into tb_sample (id, name, age) values(@id, @name, @age)"
$com_select = psmssql_createsql $con "select id, name, age from tb_sample where id >= @id"
$com_update = psmssql_createsql $con "update tb_sample set name = @name where id = @id"

#
# テーブルにデータ追加１
#
$transaction = psmssql_begin $con
psmssql_settran $com_insert $transaction
psmssql_clearbindsql $com_insert
psmssql_bindsql $com_insert "@id" "0001"
psmssql_bindsql $com_insert "@name" "伊藤　太郎"
psmssql_bindsql $com_insert "@age" 20
try{
	$rowcount = psmssql_execupdatesql $com_insert
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# テーブルにデータ追加２
#
$transaction = psmssql_begin $con
psmssql_settran $com_insert $transaction
psmssql_clearbindsql $com_insert
psmssql_bindsql $com_insert "@id" "0002"
psmssql_bindsql $com_insert "@name" "伊藤　次郎"
psmssql_bindsql $com_insert "@age" 19
try{
	$rowcount = psmssql_execupdatesql $com_insert
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# テーブルにデータ追加３
#
$transaction = psmssql_begin $con
psmssql_settran $com_insert $transaction
psmssql_clearbindsql $com_insert
psmssql_bindsql $com_insert "@id" "0003"
psmssql_bindsql $com_insert "@name" "伊藤　三郎"
psmssql_bindsql $com_insert "@age" 18
try{
	$rowcount = psmssql_execupdatesql $com_insert
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# わざと重複キーでインサートして、どんなエラーメッセージが出るか観察してみる...
#
$transaction = psmssql_begin $con
psmssql_settran $com_insert $transaction
psmssql_clearbindsql $com_insert
psmssql_bindsql $com_insert "@id" "0003"
psmssql_bindsql $com_insert "@name" "伊藤　三郎"
psmssql_bindsql $com_insert "@age" 18
try{
	$rowcount = psmssql_execupdatesql $com_insert
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
#	"0" 個の引数を指定して "ExecuteNonQuery" を呼び出し中に例外が発生しました: "制約 'PK__tb_sampl__3213E83F3E6BC8CA' の PRIMARY KEY 違反。オブジェクト 'dbo.tb_sample' には重複するキーを挿入できません。重複するキーの値は (00000007) です。
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# テーブルを参照
#
echo "== 追加後 =="
$transaction = psmssql_begin $con
psmssql_settran $com_select $transaction
psmssql_clearbindsql $com_select
psmssql_bindsql $com_select "@id" "0000"
$reader = psmssql_execsql $com_select
try{
	while (psmssql_fetch $reader){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	psmssql_free $reader
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_free $reader
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# テーブルを更新
#
$transaction = psmssql_begin $con
psmssql_settran $com_update $transaction
psmssql_clearbindsql $com_update
psmssql_bindsql $com_update "@id" "0003"
$bind_name = psmssql_bindsql $com_update "@name" "伊藤　花子"
try{
	$rowcount = psmssql_execupdatesql $com_update
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# テーブルを参照
#
echo "== 更新後 =="
$transaction = psmssql_begin $con
psmssql_settran $com_select $transaction
psmssql_clearbindsql $com_select
psmssql_bindsql $com_select "@id" "0000"
$reader = psmssql_execsql $com_select
try{
	while (psmssql_fetch $reader){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	psmssql_free $reader
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_free $reader
	psmssql_rollback $transaction
}
psmssql_free $reader
psmssql_free $transaction

#
# テーブルからデータ削除
#
$transaction = psmssql_begin $con
psmssql_settran $com_delete $transaction
psmssql_clearbindsql $com_delete
psmssql_bindsql $com_delete "@id" "0000"
try{
	$rowcount = psmssql_execupdatesql $com_delete
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_rollback $transaction
}
psmssql_free $transaction

#
# テーブルを参照
#
echo "== 全件削除後 =="
$transaction = psmssql_begin $con
psmssql_settran $com_select $transaction
psmssql_clearbindsql $com_select
psmssql_bindsql $com_select "@id" "0000"
$reader = psmssql_execsql $com_select
try{
	while (psmssql_fetch $reader){
		write-output ("{0},{1},{2}" -f $reader["id"], $reader["name"], $reader["age"])
	}
	psmssql_free $reader
	psmssql_commit $transaction
}catch{
	write-output $_.Exception.Message
	psmssql_free $reader
	psmssql_rollback $transaction
}
psmssql_free $reader
psmssql_free $transaction

#
# データベースから切断
#		
psmssql_close $con
