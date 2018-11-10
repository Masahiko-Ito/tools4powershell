$port = $args[0]

$server = pssock_start $port
$shutdown = $false
while (-not $shutdown){
	$param = pssock_accept $server
	$line = pssock_readline $param
	while ($line -ne $null -and (-not $shutdown)){
		if ($line -match "^shutdown"){
			$shutdown = $true
		}
		$stat = pssock_writeline $param $line
		if ($shutdown -eq $false){
			$line = pssock_readline $param
		}
	}
	pssock_unaccept $param
}
pssock_stop $server
