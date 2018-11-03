$addr = $args[0]
$port = $args[1]
$cmd = $args[2]

$param = pssock_open $addr $port
$stat = pssock_writeline $param $cmd
$line = pssock_readline $param
write-output $line
pssock_close $param
