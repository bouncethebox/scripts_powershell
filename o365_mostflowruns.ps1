$FlowRuns = Get-AdminFlow | Select DisplayName, @{Name = "Runs last month" ; Expression = { @(Get-FlowRun -FlowName $_.FlowName).count } } | Sort-Object 'Runs last month' -Descending

$FlowRuns| Export-Csv -Path "C:\PowerShell\FlowRuns.csv" -NoTypeInformation
