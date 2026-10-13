# Input plans are validated by Project-Codex.ps1 before calling this writer.
function Write-ProjectPlanXml($Plan,[string]$Path) {
    function DurationText([double]$Minutes) {
        $s = [math]::Round($Minutes * 60)
        return ('PT{0}H{1}M{2}S' -f [math]::Floor($s/3600),[math]::Floor(($s%3600)/60),($s%60))
    }
    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Indent = $true; $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $stream = [IO.File]::Open($Path,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    $writer = [Xml.XmlWriter]::Create($stream,$settings)
    function Element([string]$Name,$Value) { $writer.WriteElementString($Name,[string]$Value) }
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('Project','http://schemas.microsoft.com/project')
        Element 'SaveVersion' 14
        Element 'Name' ([IO.Path]::GetFileNameWithoutExtension($Path))
        Element 'ScheduleFromStart' 1
        $start = if ($Plan.PSObject.Properties['startDate']) { [datetime]::Parse($Plan.startDate,[Globalization.CultureInfo]::InvariantCulture) } else { [datetime]::Today.AddHours(8) }
        Element 'StartDate' ($start.ToString('yyyy-MM-ddTHH:mm:ss'))
        Element 'CalendarUID' 1
        Element 'DefaultStartTime' '08:00:00'
        Element 'DefaultFinishTime' '17:00:00'
        Element 'MinutesPerDay' 480
        Element 'MinutesPerWeek' 2400
        Element 'DaysPerMonth' 20
        $writer.WriteStartElement('Calendars'); $writer.WriteStartElement('Calendar')
        Element 'UID' 1; Element 'Name' 'Agent Standard 8h'; Element 'IsBaseCalendar' 1; Element 'BaseCalendarUID' -1
        $writer.WriteStartElement('WeekDays')
        for ($day=1;$day -le 7;$day++) {
            $writer.WriteStartElement('WeekDay'); Element 'DayType' $day
            $working = $day -ge 2 -and $day -le 6
            Element 'DayWorking' ([int]$working)
            if ($working) {
                $writer.WriteStartElement('WorkingTimes')
                foreach ($interval in @(@('08:00:00','12:00:00'),@('13:00:00','17:00:00'))) {
                    $writer.WriteStartElement('WorkingTime'); Element 'FromTime' $interval[0]; Element 'ToTime' $interval[1]; $writer.WriteEndElement()
                }
                $writer.WriteEndElement()
            }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement(); $writer.WriteEndElement(); $writer.WriteEndElement()
        $ids = @{}; $id = 0
        foreach ($task in $Plan.tasks) { $id++; $ids[[string]$task.key] = $id }
        $writer.WriteStartElement('Tasks')
        foreach ($task in $Plan.tasks) {
            $writer.WriteStartElement('Task')
            Element 'UID' $ids[[string]$task.key]; Element 'ID' $ids[[string]$task.key]; Element 'Name' $task.name
            Element 'IsNull' 0; Element 'OutlineLevel' 1
            # Project's importer expects explicit hours, minutes and seconds (not PnD or PT8H).
            Element 'Duration' (DurationText $task.durationMinutes)
            Element 'DurationFormat' 7
            Element 'Milestone' ([int]([double]$task.durationMinutes -eq 0))
            Element 'Summary' 0
            if ($task.PSObject.Properties['percentComplete']) { Element 'PercentComplete' $task.percentComplete }
            $percent = if ($task.PSObject.Properties['percentComplete']) {[double]$task.percentComplete} else {0}
            Element 'ActualDuration' (DurationText ($task.durationMinutes * $percent / 100))
            Element 'RemainingDuration' (DurationText ($task.durationMinutes * (100-$percent) / 100))
            Element 'ConstraintType' 0; Element 'CalendarUID' -1
            if ($task.PSObject.Properties['predecessorKeys']) {
                foreach ($key in $task.predecessorKeys) {
                    $writer.WriteStartElement('PredecessorLink')
                    Element 'PredecessorUID' $ids[[string]$key]; Element 'Type' 1; Element 'CrossProject' 0; Element 'LinkLag' 0; Element 'LagFormat' 7
                    $writer.WriteEndElement()
                }
            }
            Element 'Active' 1; Element 'Manual' 0
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement(); $writer.WriteEndElement(); $writer.WriteEndDocument()
    } finally { $writer.Dispose(); $stream.Dispose() }
}
