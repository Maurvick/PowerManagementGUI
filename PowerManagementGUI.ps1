# --- 1. Admin Check ---
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (!($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Warning "This script requires Administrator privileges to modify device settings."
    Start-Sleep -Seconds 3
    Exit
}

# --- 2. Load Required Assemblies ---
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- 3. Define XAML ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Device Power Management Manager" Height="550" Width="700" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="CheckBox">
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5,0,0,0"/>
            <Setter Property="Padding" Value="10,3"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <StackPanel Grid.Column="0">
                <TextBlock Text="Power Management Controller" FontSize="18" FontWeight="Bold"/>
                <TextBlock Text="Click headers to sort. Checked = Power Saving ON." Foreground="Gray" FontSize="11" Margin="0,2,0,0"/>
            </StackPanel>

            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom">
                <Button Name="btnEnableAll" Content="Enable All" Background="#DDFFDD"/>
                <Button Name="btnDisableAll" Content="Disable All" Background="#FFDDDD"/>
            </StackPanel>
        </Grid>

        <ListView Name="lstDevices" Grid.Row="1" BorderBrush="#FFABADB3" BorderThickness="1">
            <ListView.View>
                <GridView>
                    <GridViewColumn Width="100">
                        <GridViewColumnHeader Content="Power Saving" Tag="IsPowerSaveEnabled"/>
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <CheckBox IsChecked="{Binding IsPowerSaveEnabled, Mode=TwoWay}" Tag="{Binding Id}"/>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Width="350">
                        <GridViewColumnHeader Content="Device Name" Tag="Name"/>
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <TextBlock Text="{Binding Name}"/>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <GridViewColumn Width="150">
                        <GridViewColumnHeader Content="Type" Tag="Class"/>
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <TextBlock Text="{Binding Class}"/>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                </GridView>
            </ListView.View>
        </ListView>

        <StatusBar Grid.Row="2" Margin="0,10,0,0">
            <TextBlock Name="txtStatus" Text="Ready"/>
        </StatusBar>
    </Grid>
</Window>
"@

# --- 4. Load XAML ---
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find Controls
$lstDevices   = $window.FindName("lstDevices")
$txtStatus    = $window.FindName("txtStatus")
$btnEnableAll = $window.FindName("btnEnableAll")
$btnDisableAll= $window.FindName("btnDisableAll")

# Global variables
$script:devicesList = @()
$script:lastSortCol = ""
$script:isAscending = $true

# --- 5. Scan Logic ---
function Get-PowerManagedDevices {
    $txtStatus.Text = "Scanning devices... Please wait."
    [System.Windows.Forms.Application]::DoEvents()

    $pnpDevices = Get-CimInstance -ClassName Win32_PnPEntity
    
    try {
        $powerMgmtObjs = Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -ErrorAction Stop
    }
    catch {
        $txtStatus.Text = "Error: Failed to query WMI. (Are you Admin?)"
        return
    }

    $guiList = @()

    foreach ($pmObj in $powerMgmtObjs) {
        $cleanId = $pmObj.InstanceName -replace "_0$", ""
        $match = $pnpDevices | Where-Object { $_.PNPDeviceID -eq $cleanId }
        
        if ($match) {
            $guiList += [PSCustomObject]@{
                Name               = $match.Name
                Class              = $match.PNPClass
                Id                 = $pmObj.InstanceName
                IsPowerSaveEnabled = $pmObj.Enable
                WmiObject          = $pmObj
            }
        }
    }
    
    $script:devicesList = $guiList 
    Update-ListView "Class" # Initial Sort
    $txtStatus.Text = "Loaded $($guiList.Count) devices."
}

# --- 6. Sorting Helper ---
function Update-ListView($sortBy) {
    if ($sortBy) {
        if ($sortBy -eq $script:lastSortCol) {
            $script:isAscending = -not $script:isAscending
        } else {
            $script:isAscending = $true
            $script:lastSortCol = $sortBy
        }
    }
    
    # Apply sort to the list view
    if ($script:isAscending) {
        $lstDevices.ItemsSource = $script:devicesList | Sort-Object $script:lastSortCol
    } else {
        $lstDevices.ItemsSource = $script:devicesList | Sort-Object $script:lastSortCol -Descending
    }
}

# --- 7. Bulk Update Logic ---
function Set-AllPowerSaving($enableState) {
    $actionName = if ($enableState) { "Enabling" } else { "Disabling" }
    $total = $script:devicesList.Count
    $current = 0

    foreach ($item in $script:devicesList) {
        $current++
        
        # Only update if the state is actually changing
        if ($item.IsPowerSaveEnabled -ne $enableState) {
            $txtStatus.Text = "$actionName power saving: $current of $total ($($item.Name))"
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                $item.WmiObject.Enable = $enableState
                Set-CimInstance -InputObject $item.WmiObject -ErrorAction Stop
                
                # Update local object so UI reflects it
                $item.IsPowerSaveEnabled = $enableState
            }
            catch {
                Write-Host "Failed to update $($item.Name)"
            }
        }
    }

    # Refresh the ListView to show new checkboxes
    $lstDevices.Items.Refresh()
    $txtStatus.Text = "Bulk update complete."
}

# --- 8. Event Handlers ---

# A. ListView Click Handler (Checkboxes & Sort Headers)
$ListActionBlock = {
    param($sender, $e)

    # 1. CheckBox Click
    if ($e.OriginalSource -is [System.Windows.Controls.CheckBox]) {
        $checkBox = $e.OriginalSource
        $dataItem = $checkBox.DataContext
        
        if ($dataItem) {
            $deviceName = $dataItem.Name
            $newState = $checkBox.IsChecked
            $wmiObj = $dataItem.WmiObject

            try {
                $wmiObj.Enable = $newState
                Set-CimInstance -InputObject $wmiObj -ErrorAction Stop
                
                $statusAction = if ($newState) { "ENABLED" } else { "DISABLED" }
                $txtStatus.Text = "Success: Power Saving $statusAction for '$deviceName'"
                
                # Sync local data
                $dataItem.IsPowerSaveEnabled = $newState
            }
            catch {
                $txtStatus.Text = "Error: Update failed. $($_.Exception.Message)"
                $checkBox.IsChecked = -not $newState
            }
        }
    }
    
    # 2. Header Click (Sort)
    elseif ($e.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
        $header = $e.OriginalSource
        if ($header.Tag) { Update-ListView $header.Tag }
    }
}

# B. Button Click Handlers
$btnEnableAll.Add_Click({ Set-AllPowerSaving $true })
$btnDisableAll.Add_Click({ Set-AllPowerSaving $false })

# Attach Delegate to ListView
$RoutedDelegate = [System.Windows.RoutedEventHandler]$ListActionBlock
$lstDevices.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $RoutedDelegate)

# --- 9. Run ---
$window.Add_Loaded({
    Get-PowerManagedDevices
})

$window.ShowDialog() | Out-Null