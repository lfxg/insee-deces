param(
[Parameter(Mandatory=$False)][Switch]$exact,
$a, $b, $c
)
$Script:args=""
write-host "Num Args: " $PSBoundParameters.Keys.Count
foreach ($key in $PSBoundParameters.keys) {
    $Script:args+= "`$$key=" + $PSBoundParameters["$key"] + "  "
}
write-host $Script:args
write-host "exact value : $exact"
# }
# END {}
