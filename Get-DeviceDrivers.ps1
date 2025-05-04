<#
.DESCRIPTION
    Converts the text output from PNPutil.exe to a PowerShell object.
    Creates an object for each driver with the information given from PNPutil.exe.
    While Get-PNPDevice or the CimClass Win32_PnPSignedDriver lists alot of informations,
    original .inf file name and its oem name, I was only able to get from PNPutil.

    Maybe I'm also just stupid and have not found it :D
#>

$PnpUtilOutput = pnputil /enum-devices /drivers

[System.Collections.ArrayList]$InstanceIDIndices = @()
$InstanceIDIndices.Clear()

$Index = 0
Foreach($Line in $PnpUtilOutput){

    if($Line -match '^Instance ID:'){
        $InstanceIDIndices.add($($Index)) | Out-Null
    }

    $Index++
}

$Counter = 1

$DriverObject = Foreach($Index in $InstanceIDIndices){

    $StartIndex = $Index
    $NextArrayValue = $InstanceIDIndices[$InstanceIDIndices.count - $($InstanceIDIndices.count - $Counter)]
    $EndIndex = $NextArrayValue - 1

    if($Counter -eq $InstanceIDIndices.count){
        break
    }

    $DriverInfo = $PnpUtilOutput[$StartIndex..$EndIndex]

    $InstanceID = ((($DriverInfo | Select-String -Pattern '^Instance ID:') -split ':')[-1]).trim()
    $DeviceDescription = ((($DriverInfo | Select-String -Pattern '^Device Description:') -split ':')[-1]).trim()
    $ClassName = ((($DriverInfo | Select-String -Pattern '^Class Name:') -split ':')[-1]).trim()
    $ClassGUID = ((($DriverInfo | Select-String -Pattern '^Class GUID:') -split ':')[-1]).trim()
    try{
        $ManufacturerName = ((($DriverInfo | Select-String -Pattern '^Manufacturer Name:') -split ':')[-1]).trim()
    }catch{
        $ManufacturerName = $null
    }
    $Status = ((($DriverInfo | Select-String -Pattern '^Status:') -split ':')[-1]).trim()
    try {
        $DriverName = ((($DriverInfo | Select-String -Pattern '^Driver Name:') -split ':')[-1]).trim()
    }catch {
        $DriverName = $null
    }

    # Matching Driver Info
    $MatchingDriverStirings = $DriverInfo | Select-String "  Driver Name:", "  Original Name:", "  Provider Name:", "  Class Name:", "  Class GUID:", "  Driver Version:", "  Signer Name:", "  Matching Device ID:", "  Driver Rank:", "  Driver Status:"  

    $MatchingDriverIndex = 0

    $MatchingDriverObject = Foreach($String in $MatchingDriverStirings){

        $IndexModifier = 0

        $OriginalName = $null

        if($MatchingDriverStirings[$($MatchingDriverIndex + 1)] -match '^\s+Original Name:'){

            $IndexModifier = 1

            $OriginalName = (($MatchingDriverStirings[$($MatchingDriverIndex + 1)] -split ':')[-1]).trim()

        }

        $MatchingDriverName = (($MatchingDriverStirings[$MatchingDriverIndex] -split ':')[-1]).trim()
        $MatchingProviderName = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 1)] -split ':')[-1]).trim()
        $MatchingClassName = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 2)] -split ':')[-1]).trim()
        $MatchingClassGUID = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 3)] -split ':')[-1]).trim()

        $MatchingDriverVersionStringArray = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 4)] -split ':')[-1]).trim() -split '\s'
        $MatchingDriverDate = $MatchingDriverVersionStringArray[0]
        $MatchingDriverVersion = $MatchingDriverVersionStringArray[-1]

        $MatchingSignerName = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 5)] -split ':')[-1]).trim()
        $MatchingMatchingDeviceID = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 6)] -split ':')[-1]).trim()
        $MatchingDriverRank = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 7)] -split ':')[-1]).trim()
        $MatchingDriverStatus = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 8)] -split ':')[-1]).trim()

        [PSCustomObject]@{

            DriverName = $MatchingDriverName
            OriginalName = $OriginalName
            ProviderName = $MatchingProviderName
            ClassName = $MatchingClassName
            ClassGUID = $MatchingClassGUID
            DriverDate = $MatchingDriverDate
            DriverVersion = $MatchingDriverVersion
            SignerName = $MatchingSignerName 
            MatchingDeviceID = $MatchingMatchingDeviceID
            DriverRank = $MatchingDriverRank
            DriverStatus = $MatchingDriverStatus
    
        }

        $MatchingDriverIndex = $MatchingDriverIndex + $IndexModifier + 9

        if($MatchingDriverIndex -ge $MatchingDriverStirings.count){
            break
        }   

    }

    if($null -eq $MatchingDriverObject){

        $MatchingDriverObject = [PSCustomObject]@{

            DriverName = $null
            OriginalName = $null
            ProviderName = $null
            ClassName = $null
            ClassGUID = $null
            DriverDate = $null
            DriverVersion = $null
            SignerName = $null 
            MatchingDeviceID = $null
            DriverRank = $null
            DriverStatus = $null
    
        }

    }

    [PSCustomObject]@{

        InstanceID = $InstanceID
        DeviceDescription = $DeviceDescription
        ClassName = $ClassName
        ClassGUID = $ClassGUID
        ManufacturerName = $ManufacturerName
        Status = $Status
        DriverName = $DriverName
        MatchingDrivers = $MatchingDriverObject
        
    }

    $Counter++
}

return $DriverObject 
#$DriverObject | ConvertTo-Json -Depth 10 | Out-File .\DriverObject.json -Encoding utf8 -Force
