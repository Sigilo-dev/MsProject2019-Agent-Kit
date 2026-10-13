# Read the active project in one native XML export instead of thousands of COM gets.
function Get-ProjectSnapshot {
    param([object]$Project,[string]$ScratchDirectory=(Join-Path (Get-Location).ProviderPath 'work\project-snapshots'))
    $snapshotApp = [Runtime.InteropServices.Marshal]::GetActiveObject('MSProject.Application')
    $xmlPath=$null
    try {
        $fullName=Get-ProjectProperty $Project 'FullName'
        if ((Get-ProjectProperty (Get-ProjectProperty $snapshotApp 'ActiveProject') 'FullName') -ne $fullName) { throw 'Snapshot requires the intended project to be active.' }
        $name=Get-ProjectProperty $Project 'Name'; $saved=Get-ProjectProperty $Project 'Saved'
        [void][IO.Directory]::CreateDirectory($ScratchDirectory)
        [string]$xmlPath=Join-Path $ScratchDirectory ([guid]::NewGuid().ToString('N')+'.xml')
        $m=[Type]::Missing
        if (-not (Invoke-ProjectMethod $snapshotApp 'FileSaveAs' @($xmlPath,$m,$m,$m,$m,$false,$m,$m,$m,'MSProject.XML'))) { throw 'Bulk XML snapshot export failed.' }
        if ((Get-ProjectProperty $Project 'FullName') -ne $fullName -or (Get-ProjectProperty $Project 'Saved') -ne $saved) { throw 'XML snapshot changed project session metadata.' }
        $readerSettings=[Xml.XmlReaderSettings]::new(); $readerSettings.DtdProcessing=[Xml.DtdProcessing]::Prohibit; $readerSettings.XmlResolver=$null
        $reader=[Xml.XmlReader]::Create($xmlPath,$readerSettings)
        $doc=[Xml.XmlDocument]::new(); $doc.XmlResolver=$null
        try {$doc.Load($reader)} finally {$reader.Dispose()}
        $ns=[Xml.XmlNamespaceManager]::new($doc.NameTable); $ns.AddNamespace('p','http://schemas.microsoft.com/project')
        function XmlText($Node,[string]$Field,[string]$Default='') {
            $child=$Node.SelectSingleNode('p:'+$Field,$ns)
            if ($null -eq $child) {return $Default}; return $child.InnerText
        }
        function XmlDate($Node,[string]$Field) {
            $value=XmlText $Node $Field
            if (-not $value) {return 'NA'}
            return ([datetime]::Parse($value,[Globalization.CultureInfo]::InvariantCulture)).ToString('o')
        }
        $nodes=@($doc.SelectNodes('/p:Project/p:Tasks/p:Task',$ns) | Where-Object {(XmlText $_ 'UID') -ne '0' -and (XmlText $_ 'IsNull' '0') -eq '0'})
        $idByUid=@{}
        foreach($node in $nodes) {$idByUid[[int](XmlText $node 'UID')]=[int](XmlText $node 'ID')}
        $rows=@(foreach($node in $nodes) {
            $links=@($node.SelectNodes('p:PredecessorLink',$ns) | ForEach-Object {
                [pscustomobject]@{UniqueID=[int](XmlText $_ 'PredecessorUID');Type=[int](XmlText $_ 'Type' '1');LagTenthsOfMinute=[int](XmlText $_ 'LinkLag' '0')}
            } | Sort-Object UniqueID,Type,LagTenthsOfMinute)
            $predecessors=@(foreach($link in $links) {
                $label=[string]$idByUid[$link.UniqueID]
                if ($link.Type -ne 1 -or $link.LagTenthsOfMinute -ne 0) {
                    $label+=@('FF','FS','SF','SS')[$link.Type]
                    if ($link.LagTenthsOfMinute -ne 0) {$label+=('{0:+0.########;-0.########}min' -f ($link.LagTenthsOfMinute/10.0))}
                }
                $label
            })
            [pscustomobject][ordered]@{
                ID=[int](XmlText $node 'ID');UniqueID=[int](XmlText $node 'UID');Name=XmlText $node 'Name'
                Duration=[Xml.XmlConvert]::ToTimeSpan((XmlText $node 'Duration' 'PT0H0M0S')).TotalMinutes
                Predecessors=$predecessors -join ',';PercentComplete=[int](XmlText $node 'PercentComplete' '0')
                Manual=(XmlText $node 'Manual' '0') -eq '1';Start=XmlDate $node 'Start';Finish=XmlDate $node 'Finish'
                Summary=(XmlText $node 'Summary' '0') -eq '1';Milestone=(XmlText $node 'Milestone' '0') -eq '1'
                ConstraintType=[int](XmlText $node 'ConstraintType' '0');ConstraintDate=XmlDate $node 'ConstraintDate'
                Estimated=(XmlText $node 'Estimated' '0') -eq '1'
                PredecessorUniqueIDs=@($links | ForEach-Object {$_.UniqueID} | Sort-Object);Dependencies=$links
                Critical=(XmlText $node 'Critical' '0') -eq '1'
            }
        })
        $root=$doc.DocumentElement
        $minutes=[double](XmlText $root 'MinutesPerDay' '480')
        return [pscustomobject]@{Name=$name;FullName=$fullName;Saved=$saved;Start=([datetime]::Parse((XmlText $root 'StartDate'),[Globalization.CultureInfo]::InvariantCulture)).ToString();HoursPerDay=$minutes/60;MinutesPerDayConversion=$minutes;TaskCount=$rows.Count;Tasks=$rows}
    } finally {
        # Only this uniquely named intermediate XML is removed; user outputs are preserved.
        if ($xmlPath -and [IO.File]::Exists($xmlPath)) {[IO.File]::Delete($xmlPath)}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($snapshotApp)
    }
}
