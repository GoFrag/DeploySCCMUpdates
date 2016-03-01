<#
    TODO:
        DONE - Fix when status bar updates so it doesn't say all deployments created until all PSJobs have completed
        DONE - Add menu option for refreshing collections
        DONE - Add menu option for refreshing software update groups
        Test that reloading the options file actually updates all the correct properties
        NOT SURE IF NEEDED - Don't allow duplicate deployment names to be created - could append timestamp to deployment name
        Add logging
        DONE - Add support for UTC time
        DONE - Allow specifying the restart setting per deployment
#>

function New-DialogBox {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$true)]
        [string]$Title,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Asterisk', 'Error', 'Exclamation', 'Hand', 'Information', 'None', 'Question', 'Stop', 'Warning')]
        [string]$MessageBoxIcon,
        [Parameter(Mandatory=$true)]
        [ValidateSet('AbortRetryIgnore', 'OK', 'OKCancel', 'RetryCancel', 'YesNo', 'YesNoCancel')]
        [string]$MessageBoxButtons = 'OK'
    )

    $Icon = [System.Windows.Forms.MessageBoxIcon]::$MessageBoxIcon
    
    return [System.Windows.Forms.MessageBox]::Show($Message, $Title, $MessageBoxButtons, $Icon)
}

function Remove-DeploymentSchedule {
    param (
        [Parameter(Mandatory=$true)]
            $Control
    )

    $Control.Children.Clear()
    $Parent = $Control.Parent
    $Parent.Remove($Control)
}

function Set-HorizontalAlignment {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet('Center', 'Left', 'Right', 'Stretch')]
            [string]$HorizontalAlignment
    )

    switch ($HorizontalAlignment) {
        'Center' { return [System.Windows.HorizontalAlignment]::Center }
        'Left' { return [System.Windows.HorizontalAlignment]::Left }
        'Right' { return [System.Windows.HorizontalAlignment]::Right }
        'Stretch' { return [System.Windows.HorizontalAlignment]::Stretch }
    } 
}

function New-GroupBox {
    param (
        [Parameter(Mandatory=$true)]
            $Name,
        [Parameter(Mandatory=$true)]
            [string]$Header,
        [Parameter(Mandatory=$true)]
            $Margin,
        [Parameter(Mandatory=$true)]
            $ParentControl
    )

    $Temp = New-Object -TypeName System.Windows.Controls.GroupBox
    $Temp.Name = $Name.ToString()
    $Temp.Header = $Header
    $Temp.Margin = $Margin

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
}

function New-StackPanel {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$true)]
            $ParentControl    
    )

    $Temp = New-Object -TypeName System.Windows.Controls.StackPanel
    $Temp.Name = $Name.ToString()

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
}

function New-Label {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$true)]
            [string]$Content,
        [Parameter(Mandatory=$true)]
            [int]$Row,
        [Parameter(Mandatory=$true)]
            [int]$Column,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Center', 'Left', 'Right', 'Stretch')]
            [string]$HorizontalAlignment = 'Center',
        [Parameter(Mandatory=$true)]
            $ParentControl    
    )

    $Temp = New-Object -TypeName System.Windows.Controls.Label
    $Temp.Name = $Name.ToString()
    $Temp.Content = $Content
    $Temp.HorizontalAlignment = Set-HorizontalAlignment -HorizontalAlignment $HorizontalAlignment
    
    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
    [System.Windows.Controls.Grid]::SetColumn((Get-Variable -Name $Name -ValueOnly), $Column)
    [System.Windows.Controls.Grid]::SetRow((Get-Variable -Name $Name -ValueOnly), $Row)
}

function New-ComboBox {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$true)]
            [int]$Column,
        [Parameter(Mandatory=$true)]
            [int]$Row,
        [Parameter(Mandatory=$true)]
            $ItemsSource,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Center', 'Left', 'Right', 'Stretch')]
            [string]$HorizontalAlignment = 'Center',
        [Parameter(Mandatory=$true)]
            $ParentControl,
        [Parameter(Mandatory=$false)]
            [int]$MinWidth,
        [Parameter(Mandatory=$false)]
            [int]$SelectedIndex,
        [Parameter(Mandatory=$false)]
            [string]$Margin = '0,0,0,0'
    )

    $Temp = New-Object -TypeName System.Windows.Controls.ComboBox
    $Temp.Name = $Name.ToString()
    $Temp.ItemsSource = $ItemsSource
    $Temp.HorizontalAlignment = Set-HorizontalAlignment -HorizontalAlignment $HorizontalAlignment
    $Temp.SelectedIndex = $SelectedIndex
    $Temp.Margin = $Margin
    if ($SelectedIndex -eq -1) {
        $Temp.Text = ''
    }

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
    [System.Windows.Controls.Grid]::SetColumn((Get-Variable -Name $Name -ValueOnly), $Column)
    [System.Windows.Controls.Grid]::SetRow((Get-Variable -Name $Name -ValueOnly), $Row)
}
#region A
function New-DateTimePicker {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$true)]
            [bool]$ShowCheckBox,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Long', 'Short', 'Time', 'Custom')]
            [string]$Format,
        [Parameter(Mandatory=$false)]
            [string]$CustomFormat = 'MM/dd/yyyy hh:mm tt',
        [Parameter(Mandatory=$true)]
            [int]$Column,
        [Parameter(Mandatory=$true)]
            [int]$Row,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Center', 'Left', 'Right', 'Stretch')]
            [string]$HorizontalAlignment = 'Center',
        [Parameter(Mandatory=$true)]
            $ParentControl
    )

    $Temp = New-Object Loya.Dameer.Dameer
    $Temp.Name = $Name.ToString()
    $Temp.ShowCheckBox = $ShowCheckBox
    $Temp.Format = $Format
    $Temp.CustomFormat = $CustomFormat
    $Temp.HorizontalAlignment = Set-HorizontalAlignment -HorizontalAlignment $HorizontalAlignment

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
    [System.Windows.Controls.Grid]::SetColumn((Get-Variable -Name $Name -ValueOnly), $Column)
    [System.Windows.Controls.Grid]::SetRow((Get-Variable -Name $Name -ValueOnly), $Row)
}

function New-ProgressBar {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$true)]
            [int]$Column,
        [Parameter(Mandatory=$true)]
            [int]$Row,
        [Parameter(Mandatory=$true)]
            $ParentControl,
        [Parameter(Mandatory=$true)]
            [bool]$IsIndeterminate,
        [Parameter(Mandatory=$true)]
            [int]$Width
    )

    $Temp = New-Object -TypeName System.Windows.Controls.ProgressBar
    $Temp.Name = $Name.ToString()
    $Temp.Width = $Width
    $Temp.Height = 20
    $Temp.Margin = '5,0,0,0'
    $Temp.IsIndeterminate = $true
    $Temp.Visibility = [System.Windows.Visibility]::Collapsed
    $Temp.HorizontalAlignment = 'Right'

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
    [System.Windows.Controls.Grid]::SetColumn((Get-Variable -Name $Name -ValueOnly), $Column)
    [System.Windows.Controls.Grid]::SetRow((Get-Variable -Name $Name -ValueOnly), $Row)
}

function New-Grid {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$false)]
            [int]$Rows,
        [Parameter(Mandatory=$false)]
            [int]$Columns,
        [Parameter(Mandatory=$true)]
            $ParentControl
    )

    $Temp = New-Object System.Windows.Controls.Grid
    $Temp.Name = $Name
    
    foreach ($n in 1..$Rows) {
        New-Variable -Name RowDefinition$n -Value $(New-Object System.Windows.Controls.RowDefinition)
        $Temp.RowDefinitions.Add((Get-Variable -Name RowDefinition$n -ValueOnly))
    }

    foreach ($n in 1..$Columns) {
        New-Variable -Name ColumnDefinition$n -Value $(New-Object System.Windows.Controls.ColumnDefinition)
        $Temp.ColumnDefinitions.Add((Get-Variable -Name ColumnDefinition$n -ValueOnly))
    }

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))
}

function New-Separator {
    param (
        [Parameter(Mandatory=$true)]
            [string]$Name,
        [Parameter(Mandatory=$true)]
            $ParentControl
    )

    $Temp = New-Object -TypeName System.Windows.Controls.Separator
    $Temp.Name = $Name.ToString()

    New-Variable -Name $Name -Value $Temp -Scope Script -PassThru

    $ParentControl.AddChild((Get-Variable -Name $Name -ValueOnly))  
}

function Load-OptionsFile {
    try {
        if (-not(Test-Path (Join-Path $PSScriptRoot options.json))) {
            New-DialogBox -Message $Error[0].Exception.Message -Title 'Options.json File Not Found' -MessageBoxIcon Error -MessageBoxButtons OK
            Exit    
        }
        $script:Json = Get-Content (Join-Path $PSScriptRoot options.json) | Out-String | ConvertFrom-Json
        $script:CMSiteCode = $script:Json.CMSiteCode
        $script:CMSiteServer = $script:Json.CMSiteServer
        $script:DeploymentNamePrefix = $script:Json.DeploymentNamePrefix
        $script:DeploymentOptions = $script:Json.DeploymentOptions
        $script:LogFile = $script:Json.LogFile
        $Form.Title = "Deploy SCCM Updates - $CMSiteCode"
    }
    catch {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Unhandled Exception' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
}

<#
function Get-SoftwareUpdateDeployments {
    
    $ExcludeProperty = @('__GENUS', '__CLASS', '__SUPERCLASS', '__DYNASTY', '__RELPATH', '__PROPERTY_COUNT', '__DERIVATION', '__SERVER', '__NAMESPACE', '__PATH')
    try {
        $SoftwareUpdateDeployments = Get-WmiObject -ComputerName $CMSiteServer -Class SMS_UpdatesAssignment -Namespace root\sms\site_$CMSiteCode | select -Property AssignmentName, Enabled, CreationTime, EnforcementDeadline -ExcludeProperty $ExcludeProperty
        $SoftwareUpdateDeployments | Add-Member -Name Deadline -MemberType ScriptProperty -Value {[System.Management.ManagementDateTimeConverter]::ToDateTime($_.EnforcementDeadline)}
        $SoftwareUpdateDeployments | Add-Member -Name Created -MemberType ScriptProperty -Value {[System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime)}
        $ListView_Deployments.ItemsSource = $SoftwareUpdateDeployments | Sort-Object -Property AssignmentName
    }
    catch [System.UnauthorizedAccessException] {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Access Denied' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
    catch {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Unhandled Exception' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
}
#>

function Get-Collections {
    try {
        Write-Output $(Get-WmiObject -Class SMS_Collection -ComputerName $CMSiteServer -Namespace root\sms\site_$CMSiteCode)
    }
    catch [System.UnauthorizedAccessException] {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Access Denied' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
    catch {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Unhandled Exception' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
}

function Get-SoftwareUpdateGroups {
    try {
        Write-Output $(Get-WmiObject -Class SMS_AuthorizationList -ComputerName $CMSiteServer -Namespace root\sms\site_$CMSiteCode)
    }
    catch [System.UnauthorizedAccessException] {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Access Denied' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
    catch {
        New-DialogBox -Message $Error[0].Exception.Message -Title 'Unhandled Exception' -MessageBoxIcon Error -MessageBoxButtons OK
        Exit
    }
}

function Set-ButtonState {
    param (
        [Parameter(Mandatory=$true)]
            [bool]$IsEnabled
    )

    $Button_Deploy.IsEnabled = $IsEnabled
    $Button_AddSchedule.IsEnabled = $IsEnabled
    $Button_RemoveSchedule.IsEnabled = $IsEnabled
}

function New-SoftwareUpdateDeployment {
    param (
        [Parameter(Mandatory=$true)]
            $DeploymentName,
        [Parameter(Mandatory=$true)]
            $ScheduleSettings
    )

    $SuppressReboot = $ScheduleSettings.SuppressReboot
    switch ($SuppressReboot) {
        'Servers' { $ScheduleSettings | Add-Member -Name RestartServer -Value $true -Type NoteProperty }
        'Workstations' { $ScheduleSettings | Add-Member -Name RestartWorkstation -Value $true -Type NoteProperty }
        'Servers & Workstations' { $ScheduleSettings | Add-Member -Name RestartWorkstation -Value $true -Type NoteProperty; $ScheduleSettings | Add-Member -Name RestartServer -Value $true -Type NoteProperty }
        'None' { $ScheduleSettings | Add-Member -Name RestartWorkstation -Value $false -Type NoteProperty; $ScheduleSettings | Add-Member -Name RestartServer -Value $false -Type NoteProperty }

    }

    $Schedule = $ScheduleSettings.Schedule
    $ScheduleSettings | Add-Member -Name DeploymentAvailableDay -Value $ScheduleSettings.SoftwareAvailableTime.ToString('D') -Type NoteProperty
    $ScheduleSettings | Add-Member -Name DeploymentAvailableTime -Value $ScheduleSettings.SoftwareAvailableTime.ToString('%H:%m') -Type NoteProperty
    $ScheduleSettings | Add-Member -Name EnforcementDeadlineDay -Value $ScheduleSettings.InstallationDeadline.ToString('D') -Type NoteProperty
    $ScheduleSettings | Add-Member -Name EnforcementDeadline -Value "$($ScheduleSettings.InstallationDeadline.ToString('D')) $($ScheduleSettings.InstallationDeadline.ToString('%H:%m'))" -Type NoteProperty
    $ScheduleSettings | Add-Member -Name DeploymentName -Value $DeploymentName -MemberType NoteProperty

    $ScheduleSettings | Add-Member -Name SoftwareUpdateGroupName -Value $ComboBox_SoftwareUpdateGroups.Text -MemberType NoteProperty
    $ScheduleSettings.PSObject.Properties.Remove('SoftwareAvailableTime')
    $ScheduleSettings.PSObject.Properties.Remove('InstallationDeadline')
    $ScheduleSettings.PSObject.Properties.Remove('Schedule')
    $ScheduleSettings.PSObject.Properties.Remove('SuppressReboot')
        
    $script:Parameters = @{}
    $Json.DeploymentOptions | Get-Member -MemberType NoteProperty | Where-Object { -not [string]::IsNullOrEmpty($Json.DeploymentOptions."$($_.Name)")} | ForEach-Object {$Parameters.Add($_.Name,$Json.DeploymentOptions."$($_.Name)")}
    $ScheduleSettings | Get-Member -MemberType NoteProperty | Where-Object { -not [string]::IsNullOrEmpty($ScheduleSettings."$($_.Name)")} | ForEach-Object {$Parameters.Add($_.Name,$ScheduleSettings."$($_.Name)")}
    (Get-Variable -Name ProgressBar_Schedule$Schedule -ValueOnly -Scope Script).Visibility = $Visible
    $global:DeploymentJob = Start-Job -Name "SCCM_CreateDeployment_$Schedule" -ScriptBlock {
        Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0, $env:SMS_ADMIN_UI_PATH.Length - 4) + 'ConfigurationManager.psd1')
        Push-Location ($using:CMSiteCode + ':')
        Start-CMSoftwareUpdateDeployment @using:Parameters
        Pop-Location
    }
    
    $MessageData = [pscustomobject]@{
        ProgressBar = (Get-Variable -Name ProgressBar_Schedule$Schedule -ValueOnly -Scope Script)
    }

    $null = Register-ObjectEvent -InputObject $DeploymentJob -EventName StateChanged -SourceIdentifier DeploymentJobEnd$Schedule -MessageData $MessageData -Action {        
        if ($Sender.State -eq 'Completed') {
            $DeploymentFailed = $false
            $Event.MessageData.ProgressBar.Visibility = $Collapsed
            Unregister-Event -SourceIdentifier $Event.SourceIdentifier

            if ($Sender.ChildJobs[0].Error -ne $null) {
                New-DialogBox -Title 'Unhandled Exception' -Message $Sender.ChildJobs[0].Error -MessageBoxIcon Error -MessageBoxButtons OK
                $DeploymentFailed = $true
            }
            
            if (@(Get-Job -State Running | where Name -Like 'SCCM*').Count -eq 0) {
                Set-ButtonState -IsEnabled $true
                $ComboBox_SoftwareUpdateGroups.IsEnabled = $true
                if ($DeploymentFailed) {
                    $Label_StatusBar.Content = 'Deployment failed'    
                } else {
                    $Label_StatusBar.Content = 'All deployments created'
                }
            }

            Remove-Job $Event.Sender
            Remove-Job -Name $Event.SourceIdentifier -Force
        }
    }
}

[xml]$XAML_Form = Get-Content -Raw (Join-Path $PSScriptRoot Main_Window.xaml)

Add-Type –AssemblyName PresentationFramework
Add-Type –AssemblyName PresentationCore
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Path (Join-Path $PSScriptRoot Loya.Dameer.dll)
[System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null

try {
    $XML_Node_Reader = (New-Object System.Xml.XmlNodeReader $XAML_Form)
    $Form = [Windows.Markup.XamlReader]::Load($XML_Node_Reader)
}
catch {
    New-DialogBox -Title 'Unhandled Exception' -Message $Error[0].Exception.Message -MessageBoxIcon Error -MessageBoxButtons OK
    Exit
}

$XAML_Form.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach {
    Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) -Scope Script
}

$script:DateTimePicker_SoftwareAvailableTime_Schedule1 = New-Object Loya.Dameer.Dameer
$DateTimePicker_SoftwareAvailableTime_Schedule1.ShowCheckBox = $false
$DateTimePicker_SoftwareAvailableTime_Schedule1.Format = 'Custom'
$DateTimePicker_SoftwareAvailableTime_Schedule1.CustomFormat = 'MM/dd/yyyy hh:mm tt'
$DateTimePicker_SoftwareAvailableTime_Schedule1.HorizontalAlignment = Set-HorizontalAlignment -HorizontalAlignment Center
$Grid_Schedule1.AddChild($DateTimePicker_SoftwareAvailableTime_Schedule1)
[System.Windows.Controls.Grid]::SetColumn($DateTimePicker_SoftwareAvailableTime_Schedule1, 1)
[System.Windows.Controls.Grid]::SetRow($DateTimePicker_SoftwareAvailableTime_Schedule1, 1)

$script:DateTimePicker_InstallationDeadline_Schedule1 = New-Object Loya.Dameer.Dameer
$DateTimePicker_InstallationDeadline_Schedule1.ShowCheckBox = $false
$DateTimePicker_InstallationDeadline_Schedule1.Format = 'Custom'
$DateTimePicker_InstallationDeadline_Schedule1.CustomFormat = 'MM/dd/yyyy hh:mm tt'
$DateTimePicker_InstallationDeadline_Schedule1.HorizontalAlignment = Set-HorizontalAlignment -HorizontalAlignment Center
$Grid_Schedule1.AddChild($DateTimePicker_InstallationDeadline_Schedule1)
[System.Windows.Controls.Grid]::SetColumn($DateTimePicker_InstallationDeadline_Schedule1, 2)
[System.Windows.Controls.Grid]::SetRow($DateTimePicker_InstallationDeadline_Schedule1, 1)

$script:Hidden = [System.Windows.Visibility]::Hidden
$script:Collapsed = [System.Windows.Visibility]::Collapsed
$script:Visible = [System.Windows.Visibility]::Visible

$Label_StatusBar.Style = $null

$ComboBox_AllowRestart_Schedule1.ItemsSource = @('None', 'Servers', 'Workstations', 'Servers & Workstations')
$ComboBox_TimeBasedOn_Schedule1.ItemsSource = @('LocalTime', 'UTC')
$ComboBox_AllowRestart_Schedule1.SelectedIndex = 0
$ComboBox_TimeBasedOn_Schedule1.SelectedIndex = 0

Load-OptionsFile

#####################################################################
# Event handlers
#####################################################################
$Form.Add_Loaded({
    $script:Timer = New-Object System.Windows.Threading.DispatcherTimer
    $Timer.Interval = [timespan]'0:0:1.00'
    $Timer.Add_Tick({
        # Had to put this in here or the UI thread would block the updating
        # of the other UI components until the main form was closed
        Write-Host $null
        [Windows.Input.InputEventHandler]{ $Form.UpdateLayout() }
        [Windows.Input.InputEventHandler]{ $Label_StatusBar.UpdateLayout() }
        [Windows.Input.InputEventHandler]{ $ComboBox_SoftwareUpdateGroups.UpdateLayout() }
    })

    $Timer.Start()
    $Form.Title = "Deploy SCCM Updates - $CMSiteCode"
})

$Form.Add_Closing({
    Pop-Location
    Get-Job | Remove-Job -Force
    $Timer.Stop()
})
#endregion A
$Button_AddSchedule.Add_Click({
    try {
        $script:Schedules += 1
        New-GroupBox -Name GroupBox_Schedule$script:Schedules -Header "Schedule $script:Schedules" -Margin 5 -ParentControl $StackPanel
        New-StackPanel -Name StackPanel_Schedule$script:Schedules -ParentControl $(Get-Variable -Name GroupBox_Schedule$script:Schedules -ValueOnly)
        New-Grid -Name Grid_Schedule$script:Schedules -ParentControl $(Get-Variable -Name StackPanel_Schedule$script:Schedules -ValueOnly) -Rows 2 -Columns 6
        New-Label -Name Label_Collection_Schedule$script:Schedules -Content 'Collection' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 0 -Column 0 -HorizontalAlignment Center
        New-Label -Name Label_SoftwareAvailableTime_Schedule$script:Schedules -Content 'Software Available Time' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 0 -Column 1 -HorizontalAlignment Center
        New-Label -Name Label_InstallationDeadline_Schedule$script:Schedules -Content 'Installation Deadline' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 0 -Column 2 -HorizontalAlignment Center
        New-ComboBox -Name ComboBox_Collection_Schedule$script:Schedules -ItemsSource ($script:Collections.Name | Sort-Object) -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 1 -Column 0 -HorizontalAlignment Stretch -SelectedIndex -1
        New-DateTimePicker -Name DateTimePicker_SoftwareAvailableTime_Schedule$script:Schedules -ShowCheckBox $false -Format Custom -CustomFormat 'MM/dd/yyyy hh:mm tt' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 1 -Column 1 -HorizontalAlignment Center
        New-DateTimePicker -Name DateTimePicker_InstallationDeadline_Schedule$script:Schedules -ShowCheckBox $false -Format Custom -CustomFormat 'MM/dd/yyyy hh:mm tt' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 1 -Column 2 -HorizontalAlignment Center
        New-ProgressBar -Name ProgressBar_Schedule$script:Schedules -IsIndeterminate $true -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 1 -Column 5 -Width 75
        New-Label -Name Label_AllowRestart_Schedule$script:Schedules -Content 'Suppress Reboot' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 0 -Column 3 -HorizontalAlignment Center
        New-Label -Name Label_TimeBasedOn_Schedule$script:Schedules -Content 'Time Setting' -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -Row 0 -Column 4 -HorizontalAlignment Center
        New-ComboBox -Name ComboBox_AllowRestart_Schedule$script:Schedules -ItemsSource @('None', 'Servers', 'Workstations', 'Servers & Workstations') -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -SelectedIndex 0 -Row 1 -Column 3 -HorizontalAlignment Stretch -Margin '5,0,0,0'
        New-ComboBox -Name ComboBox_TimeBasedOn_Schedule$script:Schedules -ItemsSource @('LocalTime', 'UTC') -ParentControl $(Get-Variable -Name Grid_Schedule$script:Schedules -ValueOnly) -SelectedIndex 0 -Row 1 -Column 4 -HorizontalAlignment Stretch -Margin '5,0,0,0'
        $Label_StatusBar.Content = "Schedule $script:Schedules added"
    }
    catch {
        New-DialogBox -Title 'Unhandled Exception' -Message $Error[0].Exception.Message -MessageBoxIcon Error -MessageBoxButtons OK
    }
})

$Button_RemoveSchedule.Add_Click({
    if ($StackPanel.Children.Count -ge 5) {
        $StackPanel.Children.Remove(($StackPanel.Children | select -Last 1))
        $ControlsToRemove = (
                        "GroupBox_Schedule$script:Schedules",
                        "StackPanel_Schedule$script:Schedules",
                        "Grid_Schedule$script:Schedules",
                        "Label_Collection_Schedule$script:Schedules",
                        "Label_SoftwareAvailableTime_Schedule$script:Schedules",
                        "Label_InstallationDeadline_Schedule$script:Schedules",
                        "ComboBox_Collection_Schedule$script:Schedules",
                        "DateTimePicker_SoftwareAvailableTime_Schedule$script:Schedules",
                        "DateTimePicker_InstallationDeadline_Schedule$script:Schedules",
                        "ProgressBar_Schedule$script:Schedules",
                        "ComboBox_TimeBasedOn_Schedule$script:Schedules",
                        "ComboBox_AllowRestart_Schedule$script:Schedules",
                        "Label_TimeBasedOn_Schedule$script:Schedules",
                        "Label_AllowRestart_Schedule$script:Schedules"

        )

        foreach ($Control in $ControlsToRemove) {
            Remove-Variable -Name $Control -Scope Script
        }

        $Label_StatusBar.Content = "Schedule $script:Schedules removed"
        $script:Schedules -= 1
    } else {
        $Label_StatusBar.Content = "There must be at least 1 schedule"
        $script:Schedules = 1
    }
})

$Button_Deploy.Add_Click({
    $OkayToDeploy = New-DialogBox -Title 'Proceed?' -Message 'Are you sure you want to create these deployments?' -MessageBoxIcon Question -MessageBoxButtons YesNo

    if ($OkayToDeploy -eq 'Yes') {
        try {
            Set-ButtonState -IsEnabled $false
            $ComboBox_SoftwareUpdateGroups.IsEnabled = $false
            $NewDeployments = @()

            foreach ($Schedule in 1..$script:Schedules) {
                $obj = [pscustomobject]@{
                    'CollectionName'        = (Get-Variable -Name ComboBox_Collection_Schedule$Schedule -ValueOnly).SelectedItem
                    'SoftwareAvailableTime' = (Get-Variable DateTimePicker_SoftwareAvailableTime_Schedule$Schedule -ValueOnly).Value
                    'InstallationDeadline'  = (Get-Variable -Name DateTimePicker_InstallationDeadline_Schedule$Schedule -ValueOnly).Value
                    'SuppressReboot'          = (Get-Variable -Name ComboBox_AllowRestart_Schedule$Schedule -ValueOnly).SelectedItem
                    'TimeBasedOn'           = if ((Get-Variable -Name ComboBox_TimeBasedOn_Schedule$Schedule -ValueOnly).SelectedItem -eq 'LocalTime') {
                                                [Microsoft.ConfigurationManagement.Cmdlets.Deployments.Commands.TimeType]::LocalTime
                                              } else {
                                                [Microsoft.ConfigurationManagement.Cmdlets.Deployments.Commands.TimeType]::UTC
                                              }
                    'Schedule'              = $Schedule
                }

                $NewDeployments += $obj
            }

            foreach ($DeploymentSettings in $NewDeployments) {
                $DeploymentName = "$DeploymentNamePrefix - $($DeploymentSettings.CollectionName)"
                New-SoftwareUpdateDeployment -DeploymentName $DeploymentName -ScheduleSettings $DeploymentSettings
                $Label_StatusBar.Content = "Creating deployments. Please wait . . ."
            }
        }
        catch {
            New-DialogBox -Title 'Unhandled Exception' -Message $Error[0].Exception.Message -MessageBoxIcon Error -MessageBoxButtons OK
        }
    }
})

$MenuItem_Exit.Add_Click({
    Pop-Location
    $Form.Close()
})

$MenuItem_EditOptions.Add_Click({
    & (Join-Path $PSScriptRoot options.json)
})

$MenuItem_ReloadOptions.Add_Click({
    Load-OptionsFile
})

$MenuItem_RefreshSoftwareUpdateGroups.Add_Click({
    Get-SoftwareUpdateGroups
    $ComboBox_SoftwareUpdateGroups.ItemsSource = $SoftwareUpdateGroups | select -ExpandProperty LocalizedDisplayName | Sort-Object
})

$MenuItem_RefreshCollections.Add_Click({
    $script:Collections = Get-Collections
    
    foreach ($i in 1..$Schedules) {        $ComboBox = (Get-Variable -Name ComboBox_Collection_Schedule$i -ValueOnly -Scope Script)        $ComboBox.ItemsSource = $script:Collections.Name | Sort-Object    }
})

try {
    Import-Module ($env:SMS_ADMIN_UI_PATH.Substring(0, $env:SMS_ADMIN_UI_PATH.Length - 4) + 'ConfigurationManager.psd1')
    $script:Schedules = 1
    Write-Host -Object "Loading data for site $($CMSiteCode.ToUpper()) from $($CMSiteServer.ToUpper()) . . ."
    $SoftwareUpdateGroups = Get-SoftwareUpdateGroups
    $ComboBox_SoftwareUpdateGroups.ItemsSource = $SoftwareUpdateGroups | select -ExpandProperty LocalizedDisplayName | Sort-Object
    $script:Collections = Get-Collections
    $ComboBox_Collection_Schedule1.ItemsSource = $script:Collections.Name | Sort-Object
    #Get-SoftwareUpdateDeployments
    $Form.ShowDialog() | Out-Null
}
catch {
    New-DialogBox -Title 'Unhandled Exception' -Message $Error[0].Exception.Message -MessageBoxIcon Error -MessageBoxButtons OK
}