Param (
    [Parameter(ValueFromPipeline=$True)]
    $user
)

Process {
    $user | Set-ADUser -Replace @{pwdlastset=0}
    $user | Set-ADUser -Replace @{pwdlastset=-1}
}