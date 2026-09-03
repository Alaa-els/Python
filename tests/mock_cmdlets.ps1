# Mock Windows cmdlets so the collector can run under pwsh on Linux.
$script:SampleTick = 0
function New-Obj { param([hashtable]$Props) return [pscustomobject]$Props }
function Get-CimInstance {
    [CmdletBinding()]
    param([Parameter(Position=0)][string]$ClassName, [string]$Namespace = 'root/cimv2', [string]$Filter)
    switch ($ClassName) {
        'Win32_ComputerSystem' { return New-Obj @{ Manufacturer='LENOVO'; Model='21CB0000XX'; SystemFamily='ThinkPad X1'; SystemType='x64-based PC'; PCSystemType=2; TotalPhysicalMemory=[uint64]17179869184; NumberOfLogicalProcessors=8 } }
        'Win32_OperatingSystem' { return New-Obj @{ Caption='Microsoft Windows 11 Pro'; Version='10.0.26100'; BuildNumber='26100'; OSArchitecture='64-bit'; LastBootUpTime=(Get-Date).AddDays(-16); InstallDate=(Get-Date).AddDays(-400); TotalVisibleMemorySize=16777216; FreePhysicalMemory=4000000; TotalVirtualMemorySize=33554432; FreeVirtualMemory=12000000 } }
        'Win32_BIOS' { return New-Obj @{ SMBIOSBIOSVersion='N3AET70W'; ReleaseDate=(Get-Date).AddDays(-200); SerialNumber='PF1ABCDE' } }
        'Win32_ComputerSystemProduct' { return New-Obj @{ UUID='12345678-1234-1234-1234-123456789ABC' } }
        'Win32_Processor' { return New-Obj @{ Name='12th Gen Intel(R) Core(TM) i7-1260P'; NumberOfCores=12; NumberOfLogicalProcessors=16; MaxClockSpeed=2100; CurrentClockSpeed=2100; L2CacheSize=9216; L3CacheSize=18432; SocketDesignation='U3E1'; VirtualizationFirmwareEnabled=$true; LoadPercentage=7 } }
        'Win32_PhysicalMemory' { return @((New-Obj @{ DeviceLocator='Controller0-ChannelA'; BankLabel='BANK 0'; Capacity=[uint64]8589934592; Speed=5200; ConfiguredClockSpeed=5200; Manufacturer='Samsung'; PartNumber='M425R1GB4BB0-CQKOD '; SMBIOSMemoryType=34; FormFactor=12 }), (New-Obj @{ DeviceLocator='Controller1-ChannelA'; BankLabel='BANK 1'; Capacity=[uint64]8589934592; Speed=5200; ConfiguredClockSpeed=5200; Manufacturer='Samsung'; PartNumber='M425R1GB4BB0-CQKOD'; SMBIOSMemoryType=34; FormFactor=12 })) }
        'Win32_PhysicalMemoryArray' { return New-Obj @{ MemoryDevices=2; MaxCapacity=67108864; MaxCapacityEx=67108864 } }
        'Win32_PageFileUsage' { return New-Obj @{ Name='C:\pagefile.sys'; AllocatedBaseSize=2432; CurrentUsage=300; PeakUsage=900 } }
        'Win32_VideoController' { return New-Obj @{ Name='Intel(R) Iris(R) Xe Graphics'; AdapterRAM=1073741824; DriverVersion='31.0.101.5186'; DriverDate=(Get-Date).AddDays(-90); CurrentHorizontalResolution=1920; CurrentVerticalResolution=1200; CurrentRefreshRate=60; PNPDeviceID='PCI\VEN_8086&DEV_46A6&SUBSYS_22F317AA&REV_0C\3&11583659&0&10'; Status='OK'; VideoProcessor='Intel(R) Iris(R) Xe Graphics Family' } }
        'Win32_LogicalDisk' { return New-Obj @{ DeviceID='C:'; VolumeName='Windows'; FileSystem='NTFS'; Size=[uint64]1023000000000; FreeSpace=[uint64]90000000000 } }
        'Win32_Battery' { return New-Obj @{ BatteryStatus=2; Status='OK'; EstimatedChargeRemaining=88; EstimatedRunTime=71582788; Chemistry=2 } }
        'BatteryStaticData' { return New-Obj @{ DesignedCapacity=57000 } }
        'BatteryFullChargedCapacity' { return New-Obj @{ FullChargedCapacity=41000 } }
        'BatteryCycleCount' { throw 'Invalid class' }
        'MSAcpi_ThermalZoneTemperature' { throw 'Access denied' }
        'Win32_PnPEntity' { return @() }
        'AntivirusProduct' { return New-Obj @{ displayName='Windows Defender'; productState=397568; timestamp='Thu, 01 Jan 2026 00:00:00 GMT' } }
        'Win32_StartupCommand' { return New-Obj @{ Name='OneDrive'; Command='"C:\Users\x\AppData\Local\Microsoft\OneDrive\OneDrive.exe" /background'; Location='HKU\S-1-5-21\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'; User='X\alaa' } }
        'Win32_Service' { return New-Obj @{ DisplayName='Software Protection'; Name='sppsvc'; State='Stopped'; StartMode='Auto'; DelayedAutoStart=$true } }
        'Win32_ReliabilityRecords' { return @(1..12 | ForEach-Object { New-Obj @{ TimeGenerated=(Get-Date).AddDays(-$_); SourceName='Application Error'; ProductName='EXCEL.EXE'; EventIdentifier=1000 } }) }
        'Win32_PerfFormattedData_PerfOS_Processor' {
            $script:SampleTick++
            $load = if ($script:SampleTick % 2 -eq 0) { 85 } else { 12 }
            return @((New-Obj @{ Name='_Total'; PercentProcessorTime=$load }), (New-Obj @{ Name='0'; PercentProcessorTime=95 }), (New-Obj @{ Name='1'; PercentProcessorTime=10 }))
        }
        'Win32_PerfFormattedData_Counters_ProcessorInformation' {
            # High frequency under load, low frequency idle: must NOT flag throttling.
            $perf = if ($script:SampleTick % 2 -eq 0) { 130 } else { 45 }
            return New-Obj @{ Name='_Total'; PercentProcessorPerformance=$perf; PercentofMaximumFrequency=100; ProcessorFrequency=2100; PercentProcessorTime=50 }
        }
        'Win32_PerfFormattedData_PerfOS_Memory' { return New-Obj @{ AvailableMBytes=2600; PercentCommittedBytesInUse=61; CommittedBytes=[uint64]20000000000; CommitLimit=[uint64]34000000000; PagesInputPersec=320; PagesPersec=340; PageFaultsPersec=5000; CacheBytes=[uint64]900000000 } }
        'Win32_PerfFormattedData_PerfDisk_PhysicalDisk' { return @((New-Obj @{ Name='_Total'; PercentDiskTime=97; PercentIdleTime=3; CurrentDiskQueueLength=3; AvgDiskQueueLength=2.5; DiskReadBytesPersec=52428800; DiskWriteBytesPersec=10485760; DiskTransfersPersec=800; SplitIOPerSec=2 }), (New-Obj @{ Name='0 C:'; PercentDiskTime=97; PercentIdleTime=3; CurrentDiskQueueLength=3; AvgDiskQueueLength=2.5; DiskReadBytesPersec=52428800; DiskWriteBytesPersec=10485760; DiskTransfersPersec=800; SplitIOPerSec=2 })) }
        'Win32_PerfRawData_PerfDisk_PhysicalDisk' {
            $t = $script:SampleTick
            # 800 transfers per tick, each 40 ms => counter grows 800*0.04*10MHz = 320000 ticks
            return @((New-Obj @{ Name='_Total'; AvgDisksecPerTransfer=[uint64](320000000*$t); AvgDisksecPerTransfer_Base=[uint32](800*$t); AvgDisksecPerRead=[uint64](200000000*$t); AvgDisksecPerRead_Base=[uint32](500*$t); AvgDisksecPerWrite=[uint64](120000000*$t); AvgDisksecPerWrite_Base=[uint32](300*$t); Frequency_PerfTime=[uint64]10000000; Timestamp_PerfTime=[uint64](10000000*$t) }), (New-Obj @{ Name='0 C:'; AvgDisksecPerTransfer=[uint64](320000000*$t); AvgDisksecPerTransfer_Base=[uint32](800*$t); AvgDisksecPerRead=[uint64](200000000*$t); AvgDisksecPerRead_Base=[uint32](500*$t); AvgDisksecPerWrite=[uint64](120000000*$t); AvgDisksecPerWrite_Base=[uint32](300*$t); Frequency_PerfTime=[uint64]10000000; Timestamp_PerfTime=[uint64](10000000*$t) }))
        }
        'Win32_PerfFormattedData_PerfProc_Process' { return @((New-Obj @{ Name='_Total'; IDProcess=0; PercentProcessorTime=800; WorkingSetPrivate=0; IOReadBytesPersec=0; IOWriteBytesPersec=0 }), (New-Obj @{ Name='Idle'; IDProcess=0; PercentProcessorTime=700; WorkingSetPrivate=0; IOReadBytesPersec=0; IOWriteBytesPersec=0 }), (New-Obj @{ Name='msedge'; IDProcess=100; PercentProcessorTime=80; WorkingSetPrivate=[uint64]314572800; IOReadBytesPersec=1048576; IOWriteBytesPersec=0 }), (New-Obj @{ Name='msedge#3'; IDProcess=101; PercentProcessorTime=160; WorkingSetPrivate=[uint64]209715200; IOReadBytesPersec=0; IOWriteBytesPersec=0 }), (New-Obj @{ Name='EXCEL'; IDProcess=200; PercentProcessorTime=400; WorkingSetPrivate=[uint64]2147483648; IOReadBytesPersec=0; IOWriteBytesPersec=0 })) }
        default { throw "Invalid class $ClassName" }
    }
}
function Get-ItemProperty {
    [CmdletBinding()] param([Parameter(Position=0)][string]$Path, [string]$LiteralPath)
    $p = if ($Path) { $Path } else { $LiteralPath }
    if ($p -match 'CurrentVersion$') { return New-Obj @{ DisplayVersion='24H2'; UBR=4351 } }
    if ($p -match 'Session Manager') { return New-Obj @{ PendingFileRenameOperations=@('\??\C:\x') } }
    if ($p -match '0000') { return New-Obj @{ DriverDesc='Intel(R) Iris(R) Xe Graphics'; MatchingDeviceId='pci\ven_8086&dev_46a6&subsys_22f317aa'; 'HardwareInformation.qwMemorySize'=[int64]8589934592 } }
    if ($p -match 'Uninstall') { return @((New-Obj @{ DisplayName='Autodesk Civil 3D 2025'; DisplayVersion='13.7'; Publisher='Autodesk'; InstallDate='20250101' })) }
    throw "no mock for $p"
}
function Test-Path { [CmdletBinding()] param([Parameter(Position=0)][string]$Path, [string]$LiteralPath, [switch]$PathType)
    $p = if ($Path) { $Path } else { $LiteralPath }
    if ($p -match 'RebootPending|RebootRequired') { return $false }
    return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $p)
}
function Get-ChildItem { [CmdletBinding()] param([Parameter(Position=0)][string]$Path)
    if ($Path -match '4d36e968') { return @((New-Obj @{ PSChildName='0000'; PSPath='HKLM:\...\0000' })) }
    return @(Microsoft.PowerShell.Management\Get-ChildItem -Path $Path)
}
function Confirm-SecureBootUEFI { throw 'Access was denied' }
function Get-PhysicalDisk { return @((New-Obj @{ FriendlyName='SAMSUNG MZVL21T0HCLR'; MediaType='SSD'; BusType='NVMe'; Size=[uint64]1024209543168; HealthStatus='Healthy'; OperationalStatus=@('OK'); FirmwareVersion='GXA7601Q'; SerialNumber='S6XXX'; SpindleSpeed=0; DeviceId='0' })) }
function Get-StorageReliabilityCounter { throw 'Access denied' }
function Get-NetAdapter { param([switch]$Physical) return @((New-Obj @{ Name='Wi-Fi'; InterfaceDescription='Intel(R) Wi-Fi 6E AX211 160MHz'; Status='Up'; LinkSpeed='1.2 Gbps'; MediaType='Native 802.11'; DriverDescription='Intel(R) Wi-Fi 6E AX211 160MHz'; DriverDate='2024-03-01'; DriverVersion='23.30.0.6'; MacAddress='AA-BB' })) }
function Get-Tpm { throw 'Access denied' }
function Get-NetFirewallProfile { return @((New-Obj @{ Name='Domain'; Enabled='True'; DefaultInboundAction='Block'; DefaultOutboundAction='Allow' }), (New-Obj @{ Name='Public'; Enabled='False'; DefaultInboundAction='Block'; DefaultOutboundAction='Allow' })) }
function Get-BitLockerVolume { throw 'Access is denied' }
function Get-MpComputerStatus { return New-Obj @{ AMServiceEnabled=$true; AntivirusEnabled=$true; RealTimeProtectionEnabled=$true; AntivirusSignatureAge=12; QuickScanAge=1; FullScanAge=30; IsTamperProtected=$true; AMRunningMode='Normal' } }
function Get-HotFix { return @((New-Obj @{ HotFixID='KB5060842'; Description='Security Update'; InstalledOn=(Get-Date).AddDays(-10) })) }
function Get-WinEvent { [CmdletBinding()] param([hashtable]$FilterHashtable, [int]$MaxEvents)
    if ($FilterHashtable.ProviderName) {
        return @((New-Obj @{ ProviderName='Microsoft-Windows-Kernel-Processor-Power'; Id=37; Level=3; LevelDisplayName='Warning'; TimeCreated=(Get-Date).AddHours(-3); Message='The speed of processor 0 in group 0 is being limited by system firmware.' }), (New-Obj @{ ProviderName='disk'; Id=153; Level=3; LevelDisplayName='Warning'; TimeCreated=(Get-Date).AddHours(-5); Message='The IO operation at logical block address 0x1 for Disk 0 was retried.' }), (New-Obj @{ ProviderName='Microsoft-Windows-Kernel-Power'; Id=41; Level=1; LevelDisplayName='Critical'; TimeCreated=(Get-Date).AddDays(-2); Message='The system has rebooted without cleanly shutting down first.' }))
    }
    if ($FilterHashtable.LogName -eq 'Application') { throw 'No events were found that match the specified selection criteria.' }
    return @((New-Obj @{ ProviderName='Microsoft-Windows-Kernel-Power'; Id=41; Level=1; LevelDisplayName='Critical'; TimeCreated=(Get-Date).AddDays(-2); Message='The system has rebooted without cleanly shutting down first.' }), (New-Obj @{ ProviderName='Service Control Manager'; Id=7000; Level=2; LevelDisplayName='Error'; TimeCreated=(Get-Date).AddDays(-1); Message='The X service failed to start.' }))
}
function Get-Process { [CmdletBinding()] param() return @((New-Obj @{ ProcessName='msedge'; WorkingSet64=[int64]300000000; PrivateMemorySize64=[int64]250000000 }), (New-Obj @{ ProcessName='msedge'; WorkingSet64=[int64]200000000; PrivateMemorySize64=[int64]150000000 }), (New-Obj @{ ProcessName='EXCEL'; WorkingSet64=[int64]2200000000; PrivateMemorySize64=[int64]2100000000 })) }
