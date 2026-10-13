function Get-ProjectProperty {
    param([object]$Target, [string]$Name)
    $value = $Target.GetType().InvokeMember($Name, [Reflection.BindingFlags]::GetProperty, $null, $Target, $null)
    return ,$value
}
function Set-ProjectProperty {
    param([object]$Target, [string]$Name, [object]$Value)
    [void]$Target.GetType().InvokeMember($Name, [Reflection.BindingFlags]::SetProperty, $null, $Target, @($Value))
}
function Invoke-ProjectMethod {
    param([object]$Target, [string]$Name, [object[]]$Arguments = @())
    $flags = [Reflection.BindingFlags]::InvokeMethod -bor [Reflection.BindingFlags]::OptionalParamBinding
    $value = $Target.GetType().InvokeMember($Name, $flags, $null, $Target, $Arguments)
    return ,$value
}
function Get-ProjectSnapshot {
    param([object]$Project)
    $tasks = Get-ProjectProperty $Project 'Tasks'
    $count = [int](Get-ProjectProperty $tasks 'Count')
    $rows = @()
    for ($i=1; $i -le $count; $i++) {
        $task = $tasks.GetType().InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $tasks, @($i))
        if ($null -eq $task) { continue }
        $row = [ordered]@{}
        foreach ($name in @('ID','UniqueID','Name','Duration','Predecessors','PercentComplete','Manual','Start','Finish','Summary','Milestone','ConstraintType','ConstraintDate','Estimated')) {
            $value = Get-ProjectProperty $task $name
            if ($value -is [datetime]) { $value = $value.ToString('o') }
            $row[$name] = $value
        }
        $predecessors = Get-ProjectProperty $task 'PredecessorTasks'
        $predecessorIds = @()
        $predecessorCount = [int](Get-ProjectProperty $predecessors 'Count')
        for ($p=1; $p -le $predecessorCount; $p++) {
            $predecessor = $predecessors.GetType().InvokeMember('Item',[Reflection.BindingFlags]::GetProperty,$null,$predecessors,@($p))
            $predecessorIds += [int](Get-ProjectProperty $predecessor 'UniqueID')
        }
        $row['PredecessorUniqueIDs'] = @($predecessorIds | Sort-Object)
        $rows += [pscustomobject]$row
    }
    return [pscustomobject]@{
        Name=Get-ProjectProperty $Project 'Name'
        FullName=Get-ProjectProperty $Project 'FullName'
        Saved=Get-ProjectProperty $Project 'Saved'
        Start=[string](Get-ProjectProperty $Project 'ProjectStart')
        HoursPerDay=[double](Get-ProjectProperty $Project 'HoursPerDay')
        MinutesPerDayConversion=60 * [double](Get-ProjectProperty $Project 'HoursPerDay')
        TaskCount=$count
        Tasks=@($rows)
    }
}
