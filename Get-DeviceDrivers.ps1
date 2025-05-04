<#
.DESCRIPTION
    Converts the text output from PNPutil.exe to a PowerShell object.
    Creates an object for each driver with the information given from PNPutil.exe.
    While Get-PNPDevice or the CimClass Win32_PnPSignedDriver lists alot of informations,
    original .inf file name and its oem name, I was only able to get from PNPutil.

    Maybe I'm also just stupid and have not found it :D
#>

# Gets data an all device drivers.
$PnpUtilOutput = pnputil /enum-devices /drivers

# This array holds all line that start with Instance ID:
# That line is the start for a drivers data. All other line until the next Instance ID line, is data for that driver.
[System.Collections.ArrayList]$InstanceIDIndices = @()
$InstanceIDIndices.Clear()

#region INDEXING

# This is the line count. It is later used as index to isolate a drivers data.
# Ex.: $PnpUtilOutput[2..20]
$Index = 0

Foreach($Line in $PnpUtilOutput){

    if($Line -match '^Instance ID:'){
        $InstanceIDIndices.add($Index) | Out-Null
    }

    $Index++
}

#endregion INDEXING

#region OBJECT CREATION

<# 
This counter is used to find the end index.
The current value of the loop is the start index.

Ex.: Given a drivers data ranges from line 2 to 19 in PNPutils output, the index array holds the values 2 and 20.
But 20 is a "Instance ID:" line, so the start of anothers drivers data. The end for this drivers data is 20 - 1, 19.

Ex.: Given the index array holds 226 values, the first value being 2 and the second being 20.
We can index into the index array with $Counter being one higher then the current loop.
From this we get the value of 1th being 20, while we have the start value from the current loop variable.

To get the next drivers data, its start index value is 20, being at the 1th position in the array, so we need the 2th positions value, with $Counter being 2.
And so on...

$Counter is incremented at the end of the loop.
#>
$Counter = 1

$DriverObject = Foreach($Index in $InstanceIDIndices){

    $StartIndex = $Index
    $NextArrayValue = $InstanceIDIndices[$Counter]
    # This has to be one less of the value, or it would include the "Instance ID:" line from the next driver, which is of no use and would mess with the string selection.
    $EndIndex = $NextArrayValue - 1

    <#
    The last entry in the index array is the start of the last drivers data.
    But, there being no next driver, there is also no next "Instance ID:" line, which would mark the end of the last drivers data.
    This causes $EndIndex to be -1, which would be bad. Ex. $PnpUtilOutput[5609..-1] cause something strage to happen, almost like reversing the string array... xD
    So, the last end index is the end of the $PnpUtilOutput string array, it's total line count.
    #>
    if($Counter -eq $InstanceIDIndices.Count){
        $EndIndex = $PnpUtilOutput.count
    }
    
    # This gets a drivers data. Which are from line $StartIndex, until $EndIndex.
    $DriverInfo = $PnpUtilOutput[$StartIndex..$EndIndex]

    # Selecting the data we need. Each string is split, the last line from the resulting array is selected and trimmed.
    # The operation wrapped in a try/catch are not always there, resulting in a null-valued expression error
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

    # Grabbing the Matching Driver Info
    $MatchingDriverStirings = $DriverInfo | Select-String "  Driver Name:", "  Original Name:", "  Provider Name:", "  Class Name:", "  Class GUID:", "  Driver Version:", "  Signer Name:", "  Matching Device ID:", "  Driver Rank:", "  Driver Status:"  
    <#

    Example of what is selected:
        Driver Name:            oem69.inf
       (Original Name:          netstrUMac.inf)
        Provider Name:          Microsoft
        Class Name:             Net
        Class GUID:             {e6a4bb54-5bcc-4736-a412-f8af9484c28e}
        Driver Version:         06/21/2006 10.0.26100.1
        Signer Name:            Microsoft Windows
        Matching Device ID:     PCI\VEN_1250&DEV8896
        Driver Rank:            00FF0420
        Driver Status:          Best Ranked / Installed
    #>

    # This is used to index into the above resulting string array.
    $MatchingDriverIndex = 0

    $MatchingDriverObject = Foreach($String in $MatchingDriverStirings){

        <#
        Original Name is not always present. This results in string arrays of 9 instead of 10 strings.
        $IndexModifier is set to 1, if a "Original Name:" string is present, so that the loop can count up to 10
        with: $($MatchingDriverIndex + $IndexModifier + n), while the the max count is 9 with the $IndexModifier staying 0.
        #>
        $IndexModifier = 0
        $OriginalName = $null

        if($MatchingDriverStirings[$($MatchingDriverIndex + 1)] -match '^\s+Original Name:'){

            $IndexModifier = 1

            $OriginalName = (($MatchingDriverStirings[$($MatchingDriverIndex + 1)] -split ':')[-1]).trim()

        }

        <#
        Indexing into the array to get the matching drivers data.
        THE ORDER OF THE PROPERTIES IS IMPORTANT! The order in which they are assigned to thier variable is not, but the indexing position needs to follow PNPutil output order.

        Each string is split, the last line from the resulting array is selected and trimmed.
        #>
        $MatchingDriverName = (($MatchingDriverStirings[$MatchingDriverIndex] -split ':')[-1]).trim()
        $MatchingProviderName = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 1)] -split ':')[-1]).trim()
        $MatchingClassName = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 2)] -split ':')[-1]).trim()
        $MatchingClassGUID = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 3)] -split ':')[-1]).trim()

        # Seperating DriverVersion into the drivers date and version.
        $MatchingDriverVersionStringArray = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 4)] -split ':')[-1]).trim() -split '\s'
        $MatchingDriverDate = $MatchingDriverVersionStringArray[0]
        $MatchingDriverVersion = $MatchingDriverVersionStringArray[-1]

        $MatchingSignerName = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 5)] -split ':')[-1]).trim()
        $MatchingMatchingDeviceID = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 6)] -split ':')[-1]).trim()
        $MatchingDriverRank = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 7)] -split ':')[-1]).trim()
        $MatchingDriverStatus = (($MatchingDriverStirings[$($MatchingDriverIndex + $IndexModifier + 8)] -split ':')[-1]).trim()

        # Building the matching drivers object.
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

        <#
        IT IS IMPORTANT THAT THE INT AT THE END HERE (9) MATCHES THE AMOUNT OF SELECTED STRINGS, MINUS ONE.
        If not, indexing into the string array will not work.
        
        This sets the next start index.
        #>
        $MatchingDriverIndex = $MatchingDriverIndex + $IndexModifier + 9

        # Ends the loop. Given an array of 20 strings, once the 10th is reach, the loop can end,
        # because above, the info needed is selected from the 10th line down.
        if($MatchingDriverIndex -ge $MatchingDriverStirings.count){
            break
        }   

    }

    # This will create a null object, if a driver has no matching drivers data.
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

    # Creating the return object for each driver.
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

#end region OBJECT CREATION


return $DriverObject 
#$DriverObject | ConvertTo-Json -Depth 10 | Out-File .\DriverObject.json -Encoding utf8 -Force
