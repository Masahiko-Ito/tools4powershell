$port = $args[0]

$childFunc = {
	param($arg_param, $arg_profile)
	. $arg_profile
	$line = pssock_readline $arg_param
	while ($line -ne $null){
		$stat = pssock_writeline $arg_param $line
		$line = pssock_readline $arg_param
	}
	pssock_unaccept $arg_param
}

$refAryPowershell = psrunspc_getarraylist
$refAryChild = psrunspc_getarraylist
$runSpacePool = psrunspc_open 10

$server = pssock_start $port
while ($true){
	$param = pssock_accept $server

	$powershell = psrunspc_createthread $runSpacePool $childFunc
	psrunspc_addargument $powershell $param
	psrunspc_addargument $powershell $profile
	psrunspc_begin $powershell $refAryPowershell.value $refAryChild.value
	psrunspc_waitasync $refAryPowershell.value $refAryChild.value
}
pssock_stop $server
psrunspc_wait $refAryPowershell.value $refAryChild.value
psrunspc_close $runSpacePool
