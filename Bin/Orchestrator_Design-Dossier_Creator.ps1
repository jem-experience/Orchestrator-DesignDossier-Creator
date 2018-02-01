##############################################################################################################################################################
#                                                                                                                                                            #
#   Program : Orchestrator Design Dossier Creator                                                                                                            #
#   Developper : Jérémy GARROS following "brunosa" original script                                                                                           #
#   Contact : jeremy.garros@jem-experience.fr                                                                                                                #
#   Original Reference are on the following blog :                                                                                                           #
#     http://blogs.technet.com/b/privatecloud/archive/2014/05/08/updated-tool-smart-documentation-and-conversion-helper-for-your-orchestrator-runbooks.aspx  #
#     Windows Server and System Center Customer, Architecture and Technologies (CAT) team                                                                    #
#                                                                                                                                                            #
#   Help : Use in Parameter the Name of the Configuration File                                                                                               #
#   Version : 1.0                                                                                                                                            #
#                                                                                                                                                            #
##############################################################################################################################################################

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True)]
    [Alias("ConfigName")]
    [String]$XMLconfigFileName
    )

## Clear Error and Console ##
$Error.Clear()
Clear-Host
###################################

#################################################################################
#                                                                               #
#                                   Constants                                   #
#                                                                               #
#################################################################################

$ToolVersion = "1.0"

## If Verbose ##
If($Verbose)
{
    $VerbosePreference = "Continue"
}

## Get Configuration File ##
If ($psISE)
{
    $ScriptFilePath = $psISE.CurrentFile.FullPath
    $pathScript = $ScriptFilePath | Split-Path -Parent

    ## Use the following line to overwrite the Script Parameters for debug ##
    $XMLconfigFileName = "Orchestrator_Design-Dossier_Creator.xml"
}
Else
{
    $ScriptFilePath = $MyInvocation.MyCommand.Path
    $pathScript = [System.IO.Path]::GetDirectoryName($ScriptFilePath)
}
###################################

## Determine Path ##
If(($pathScript.EndsWith("\Bin")) -or ($pathScript.EndsWith("\Bin\")))
{
    $pathScriptParent = $pathScript | Split-Path -Parent

    $ConfigurationPath = "$pathScriptParent\Configuration\"
    $ActivitiesPicturePath = "$pathScriptParent\Orchestrator_Activities_JPG\"
    $OutputPath = "$pathScriptParent\Output\"
    Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] PathScript found! retrieving all preconfigured folders"
}
Else
{
    ## The script has been moved from Bin folder, We will try to find the Configuration file and folder ##
    $IsConfigurationFolder = Get-ChildItem -Path:$pathScript -Filter 'Configuration' -Directory
    If($IsConfigurationFolder -ne $Null)
    {
        ## Try to locate the Xml file ##
        $IsConfigurationFile = (Get-ChildItem -Path:$IsConfigurationFolder.FullName -Filter '*.xml') | Where { $_.Name -eq "$XMLconfigFileName"}
        If($IsConfigurationFile -ne $Null)
        {
            ## We will keep the Configuration Path as Origin ##
            $ConfigurationPath = "$($IsConfigurationFolder.FullName)\"
            $pathScriptParent = $ConfigurationPath | Split-Path -Parent
            $ActivitiesPicturePath = "$pathScriptParent\Orchestrator_Activities_JPG\"
            $OutputPath = "$pathScriptParent\Output\"
            Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] PathScript discovered! retrieving all preconfigured folders"
        }
        Else
        {
            Write-Host -ForegroundColor RED "["(date -format "HH:mm:ss")"] The Script was not able to find the correct path for the configuration file, please control the package"
            Exit
        }
    }
    Else
    {
        Write-Host -ForegroundColor RED "["(date -format "HH:mm:ss")"] The Script was not able to find the correct path for configuration folder, please control the package"
        Exit
    }
}

###################################

## Test the Output Path ##
If(-not(Test-Path -Path:$OutputPath))
{
    ## We will create the folder ##
    Try
    {
        New-Item -Path:$pathScriptParent -Name "Output" -ItemType Directory -Force
    }
    Catch
    {
        Write-Host -ForegroundColor Yellow "["(date -format "HH:mm:ss")"] The Script was not able to create the missing Output Folder, please create it manually or control user right"
    }
}
###################################

## Control the existency of Orchestrator picture in the Jpg folder ##
If (Test-Path -Path:$ActivitiesPicturePath)
{
    $NumberOfThumbnails = (Get-ChildItem -Path:$ActivitiesPicturePath -Filter "*.jpg").Count
    If ($NumberOfThumbnails -lt 10)
    {
        Write-Host -ForegroundColor Yellow "["(date -format "HH:mm:ss")"] WARNING : Only $NumberOfThumbnails image(s) were found in the local folder... This may mean that you haven't run the Image Export script yet (SMA-DocumentationConversionHelper-ImageExport.ps1). Visio and Word export will still work, but will use a default image."
    }
    If ((Test-Path -Path ("$ActivitiesPicturePath" + "default.jpg")) -eq $False)
    {
        Write-Host -ForegroundColor Yellow "["(date -format "HH:mm:ss")"] WARNING : 'default.jpg' thumbnail not found in the ActivitiesPicture folder. Make sure you copied it from the download package, or that you are calling the script from the same folder where this file is located. Without this file, Visio and Word export will return errors when adding thumbnails."
    }
}
Else
{
    Write-Host -ForegroundColor Orange "["(date -format "HH:mm:ss")"] Unable to find the folder containing all Orchestrator pictures generated by (SMA-DocumentationConversionHelper-ImageExport.ps1), the output document will have missing information, all Activities picture will be represented by the default picture"
}
###################################

## Get Configuration File ##
$MyConfigurationXmLFile = (Get-ChildItem -Path:$ConfigurationPath -Filter '*.xml') | Where { $_.Name -eq "$XMLconfigFileName"}
###################################

## Try to Load the XML Configuration ##
Try
{
    $XmlConfiguration = [XML](Get-Content $($MyConfigurationXmLFile.FullName))
}
Catch
{
    Write-host $Error[0]
    Exit
}
###################################

## Load Configuration ##
$DatabaseServer = $XmlConfiguration.DefaultConfiguration.DatabaseServer
$DatabasePort = $XmlConfiguration.DefaultConfiguration.DatabasePort
$DatabaseName = $XmlConfiguration.DefaultConfiguration.DatabaseName
$DatabaseUserName = $XmlConfiguration.DefaultConfiguration.DatabaseUserName
$DatabasePassword = $XmlConfiguration.DefaultConfiguration.DatabasePassword
$DatabaseSecurePassword = $XmlConfiguration.DefaultConfiguration.DatabaseSecurePassword
$Global:VisioTemplate = $XmlConfiguration.DefaultConfiguration.VisioTemplate
$Global:VisioStencil = $XmlConfiguration.DefaultConfiguration.VisioStencil
$Global:VisioCallout = $XmlConfiguration.DefaultConfiguration.VisioCallout
$Global:VisioGlueFrom = $XmlConfiguration.DefaultConfiguration.VisioGlueFrom
$Global:VisioGlueTo = $XmlConfiguration.DefaultConfiguration.VisioGlueTo
###################################


#################################################################################
#                                                                               #
#                                   Functions                                   #
#                                                                               #
#################################################################################

function ParseAndWriteProperty {
# This function takes the value of an activity property, and tries to parse published data
# and variables, to replace the data by readable content (actual names of the published data
# and variables)

    param (
    [String]$Prefix,
    [String]$TmpProperty,
    [String]$NbTab,
    [String]$ExportMode
    )

    $myConnection3 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr
    $myConnection3.Open()

    $TmpProperty=$TmpProperty.Replace("\``d.T.~Ed/", "###PUBDATA###")
    $TmpProperty=$TmpProperty.Replace("\``d.T.~Vb/", "###VAR###")
    
    If (($TmpProperty.Contains("###PUBDATA###") -eq $True) -Or ($TmpProperty.Contains("###VAR###") -eq $True))
    
            {
            #Let's work through the potential published data first, avoiding the first item in the split (not a published data)
            $ComputedProperties = $TmpProperty -split("###PUBDATA###{")
            For ($i=1; $i -lt $ComputedProperties.Length; $i++){
                #PublishedData generally has one GUIDs {activity}.publisheddataincleartext, but published data Initialize Data activities are formatted like {activity}.{parameter}
                #Let's work through the first GUID in all cases
                $ComputedProperties[$i]= ($ComputedProperties[$i] -split ("###PUBDATA###"))[0]
                $OutputID = "{" + $ComputedProperties[$i].Substring(0, 37)
                #Note : Adding a specific check to make sure OutputID is a GUID. 
                #Otherwise it fails if script has a invoke-command with published data for the computername and no explicit scriptblock option followed by {
                If ($OutputID -match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$")) {
                                $SqlQuery = "select name, ObjectType from objects where UniqueID = '" + $OutputID + "'"
                                $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                                $dr3 = $myCommand3.ExecuteReader()
                                while ($dr3.Read())
                                    {
                                    $OutputName = $dr3["name"]
                                    $OutputType = $dr3["ObjectType"].ToString
                                    If ($ExportMode -eq "DOC"){
                                        If ($Global:ActivityDependenciesActivityNames.Contains($OutputName) -eq $False) {
                                            $Global:ActivityDependenciesActivityNames += $OutputName
                                            $Global:ActivityDependenciesActivityTypes += $dr3["ObjectType"]
                                            }
                                        }
                                    }
                                $dr3.Close()
                                $TmpProperty = $TmpProperty.Replace($OutputID, "{Activity:" + $OutputName + "}")
                                #Now we can check if there is a GUID in another GUID
                                $NumberOfGUIDS = $ComputedProperties[$i].Split("{").Length - 1
                                If ($NumberOfGUIDS -eq 1) {
                                        $OutputSuffix = "{" + ($ComputedProperties[$i].Split("{")[1]).Substring(0, 36) + "}"
                                        $SqlQuery = "select value from CUSTOM_START_PARAMETERS where UniqueID = '" + $OutputSuffix + "'"
                                        $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                                        $dr3 = $myCommand3.ExecuteReader()
                                        while ($dr3.Read())
                                                {
                                                $OutputSuffixName = $dr3["value"]
                                                }
                                        $dr3.Close()
                                        $OutputSuffix = "}." + $OutputSuffix
                                        $TmpProperty = $TmpProperty.Replace($OutputSuffix, (".PublishedData:" + $OutputSuffixName + "}"))
                                    }
                                    else
                                    {
                                    $OutputSuffix = "}." + $ComputedProperties[$i].Split(".")[1]
                                    $TmpProperty = $TmpProperty.Replace($OutputSuffix, ".PublishedData:" + $ComputedProperties[$i].Split(".")[1] + "}")
                                    }

                }
            }
            #Let's work through the potential variables too, avoiding the first item in the split (not a variable)         
            $ComputedProperties = $TmpProperty -split("###VAR###{")
            For ($i=1; $i -lt $ComputedProperties.Length; $i++)
                {
                $OutputVariable = "{" + $ComputedProperties[$i].Substring(0, 37)
                If ($OutputVariable -match ("^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"))
                    {
                    $SqlQuery = "select name from objects where UniqueID = '" + $OutputVariable + "'"
                    $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                    $dr3 = $myCommand3.ExecuteReader()
                    while ($dr3.Read()) {$OutputVariableName = $dr3["name"]}
                    $dr3.Close()
                    $SqlQuery = "select value from variables where UniqueID = '" + $OutputVariable + "'"
                    $myCommand3 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection3
                    $dr3 = $myCommand3.ExecuteReader()
                    while ($dr3.Read())
                         {
                         $OutputVariableValue = $dr3["value"]
                         #Let's check if this is an encrypted or empty variable
                         If ([System.DBNull]::Value.Equals($dr3["value"]) -eq $False)
                            {$OutputVariableValue=$OutputVariableValue.Replace("\``d.T.~De/", "###ENCRYPTEDDATA###")}
                            else
                            {$OutputVariableValue="[Empty Value]"}
                         }
                    $dr3.Close()
                    If ($ExportMode -eq "DOC")
                         {
                         $Global:ActivityDependenciesVariableNames += $OutputVariableName
                         $Global:ActivityDependenciesVariableValues += $OutputVariableValue
                         }
                    $TmpProperty = $TmpProperty.Replace($OutputVariable, ("{Variable:" + $OutputVariableName + "}"))
                    $Global:FlagVariables = $True
                    If ($Global:FlagVariablesList.Contains("{" + $OutputVariableName + "}"))
                          {$Global:FlagVariablesNumber[$Global:FlagVariablesList.IndexOf("{" + $OutputVariableName + "}")] = $Global:FlagVariablesNumber[$Global:FlagVariablesList.IndexOf("{" + $OutputVariableName + "}")] +1}
                          else
                             {
                             $Global:FlagVariablesList+= "{" + $OutputVariableName + "}"
                             $Global:FlagVariablesNumber+= 1
                             If (($OutputVariableValue -ne $Null) -and ($OutputVariableValue.Contains("###ENCRYPTEDDATA###")))
                                  {$Global:FlagVariablesValue+= "[Encrypted Data]"}
                                  else
                                  {$Global:FlagVariablesValue+= $OutputVariableValue}
                             }
                    }
                }
            
            #We work on the TmpProperty, to extract the published data items
            $TmpProperty = $TmpProperty.Replace("###PUBDATA###", "####")
            $TmpProperty = $TmpProperty.Replace("###VAR###", "####")
            $TmpProperty = $TmpProperty.Replace("`r`n", "")
            $ComputedProperties = $TmpProperty -split("####")
            $FullPty = ""
            ForEach ($ComputedProperty In $ComputedProperties){
                If ($ComputedProperty -ne ""){$FullPty = $FullPty + $ComputedProperty}
            }
            $Lines = $FullPty -split("`n")
            $i = 0
            ForEach ($Line In $Lines)
                {
                If ($i -eq 0)
                    {WriteToFile -ExportMode $ExportMode -Add "$Prefix $Line" -NbTab $NbTab}
                    Else
                    {WriteToFile -ExportMode $ExportMode -Add $Line -NbTab $NbTab}
                $i = $i + 1
                }
            
            }
        else
            # There is no published data or variable in the value of this property
            {
            If ($TmpProperty.Contains("\``d.T.~De/") -eq $True) { $TmpProperty = "[Encrypted Data]"}
            WriteToFile -ExportMode $ExportMode -Add "$Prefix $TmpProperty" -NbTab $NbTab
            }
    $myConnection3.Close()
} #ParseAndWriteProperty


function WriteToFile() {
# This function is called throughout the tool, to write to the PS1 file
# For readability purposes in the PS1 file being generated, it includes
# a 'NbTab' parameter, to compute the tabulations as we go deeper
# into the PowerShell branches
    param (
    [String]$Add,
    [int]$NbTab,
    [String]$ExportMode
    )

    If($Add -match "$([char]2039) START ACTIVITY - ")
    {
        $oTable.Cell($r, 3).Range.Text = $Add
        $oTable.Cell($r, 3).Range.Paragraphs.SpaceAfter = 0
    }
    ElseIf($Add -match "$([char]2039) END ACTIVITY - ")
    {
        $oPara1 = $oTable.Cell($r, 3).Range.Paragraphs.Add()
        $oPara1.Range.Text = $Add
        $oPara1.SpaceAfter = 8
    }
    Else
    {
        $oPara1 = $oTable.Cell($r, 3).Range.Paragraphs.Add()
        $oPara1.Range.Text = $Add
        $oPara1.SpaceAfter = 0
    }
} #WriteToFile


function LinkCondition {
# This function retrieves the condition on a link betwen activities (if any)

    param (
    [String]$LinkID
    )

    $output = ""
    $OutputID = ""
    $SqlQuery = "select condition, data, value from TRIGGERS where ParentID = '" + $LinkID + "'"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
    $dr = $myCommand.ExecuteReader()
    $OffsetBracket=0
    while ($dr.Read())
            {
            $OutputID = ($dr["data"]).Substring(0, 38)
            $NumberofGUIDs = ($dr["data"]).Split("{").Count - 1
            Switch ($dr["condition"])
                {
                "isgreaterthan" {$Outputcondition = "-gt"}
                "isgreaterthanorequalto" {$Outputcondition = "-ge"}
                "islessthan" {$Outputcondition = "-lt"}
                "islessthanorequalto" {$Outputcondition = "-le"}
                "equals" {$Outputcondition = "-eq"}
                "doesnotequal" {$Outputcondition = "-ne"}
                "" {$Outputcondition = "{linkcondition:returns}" ; $OffsetBracket = 1}
                default
                    # "contains" "doesnotcontain" "endswith" "startswith" "doesnotmatchpattern" "matchespattern"
                    {
                    $Outputcondition = "{linkcondition:" + $dr["condition"] + "}"
                    $Global:FlagStringcondition = $True
                    $OffsetBracket = 1
                    }
                }
            $Output = "If (" + $dr["data"] + " " + $Outputcondition + " `"" + $dr["value"] + "`") {"
            }
    $dr.Close()
    If ($OutputID -ne "")
            {
            #There was a condition, let's convert the published data activity (first part)
            $SqlQuery = "select name, ObjectType from objects where UniqueID = '" + $OutputID + "'"
            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
            $dr = $myCommand.ExecuteReader()
            while ($dr.Read()) {$OutputName = $dr["name"]}
            $dr.Close()
            $Output = $Output.Replace($OutputID, "{Activity:" + $OutputName + "}")
            #Let's check if there is a second GUID to convert - only applicable when it's published data from an initialize data activity
            $NumberofGUIDs = $output.Split("{").Count - 1 -$OffsetBracket
            If ($NumberofGUIDs -eq 3)
                {
                $OutputSuffix = "{" + $output.Split("{")[2].Substring(0, 36) + "}"
                $SqlQuery = "select value from CUSTOM_START_PARAMETERS where UniqueID = '" + $OutputSuffix + "'"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read()) {$OutputSuffixName = $dr["value"]}
                $dr.Close()
                $OutputSuffix = "}." + $OutputSuffix
                $Output = $output.Replace($OutputSuffix, ".PublishedData" + $OutputSuffixName + "}")
                }
            }
    Return $Output
} #LinkCondition


function AppendActivityDetails {
# This function is being called to fill the details of a specific activity
# It does that in a different manner depending on the type of activity
# It also calls ParseAndWriteProperty as needed, when parsing properties values

    param (
    [String]$ActivityID,
    [String]$ActivityDetailsShort,
    [String]$ActivityType,
    [Int]$NbTab,
    [String]$ExportMode
    )

    $DoNotExportProperties = @("UniqueID", "ExecutionData", "CachedProps")
    $PubDataNames = @()
    $PubDataTypes = @()
    $PubDataValues = @()
    $XmlNames = @()
    $XmlValues = @()
    $XmlPatterns = @()
    $XmlRelations = @()

    $myConnection2 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
    $myConnection2.Open()

    #First, let's retrieve the full name of the activity type
    $SqlQuery = "select Name from OBJECTTYPES where UniqueID = '{" + $ActivityType + "}'"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
    $dr = $myCommand.ExecuteReader()
    while ($dr.Read()) {$ActivityTypeName = $dr["Name"]}
    $dr.Close()

    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Working on activity : $ActivityDetailsShort (Activity Type : $ActivityTypeName)"

    Switch ($ActivityType)
    {
    "ed7f2a41-107a-4b74-bafe-adae63632b79"
    #This is a Powershell activity
        {
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) START ACTIVITY - $ActivityDetailsShort (Activity Type : Run .NET Script)" -NbTab $NbTab
        #retrieve script and script type (Powershell, C#, VB.NET, JScript)
        $SqlQuery = "select ScriptType, ScriptBody from RUNDOTNETSCRIPT where UniqueID = '" + $ActivityID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
                {
                $TmpProperty = $dr["ScriptBody"]
                $ScriptType = $dr["ScriptType"]
                }
        $dr.Close()
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) Script Type = $ScriptType" -NbTab $NbTab
        If ($ScriptType -eq "PowerShell")
            {ParseAndWriteProperty -ExportMode $ExportMode -Prefix "" -TmpProperty $TmpProperty -NbTab $NbTab}
            else
            {ParseAndWriteProperty -ExportMode $ExportMode -Prefix "$([char]2039) " -TmpProperty $TmpProperty -NbTab $NbTab}
        #retrieve published data too
        $SqlQuery = "select publisheddata from RUNDOTNETSCRIPT where UniqueID = '" + $ActivityID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read()) {[string]$TmpProperty = $dr["publisheddata"]}
        $dr.Close()
        If ($TmpProperty -eq "")
                    {WriteToFile -ExportMode $ExportMode -Add "$([char]2039) Published Data - None" -NbTab $NbTab}
                    else
                    {
                    $PubDataNames.Clear()
                    $PubDataTypes.Clear()
                    $PubDataValues.Clear()
                    $xmlDoc = New-Object System.Xml.XmlDocument
                    $TmpProperty = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" + $TmpProperty
                    [System.Xml.XmlDocument]$xmlDoc.LoadXml($TmpProperty)
                    $Input = New-Object System.Xml.XmlNodeReader $xmlDoc
                    While ($Input.Read()){
                        If ($Input.NodeType -eq [System.Xml.XmlNodeType]::Element){
                            switch ($Input.Name){
                                "Name" {$PubDataNames+=$Input.ReadString()}
                                "Type"{$PubDataTypes+=$Input.ReadString()}
                                "Variable"{$PubDataValues+=$Input.ReadString()}
                            }
                        }
                    }
                    $Input.Close()
                    $PubDataOutput = ""
                    ForEach ($PubDataName In $PubDataNames){
                            $PubDataOutput = $PubDataOutput + " Name : " + $PubDataName + " / Type : " + $PubDataTypes.Item($PubDataNames.IndexOf($PubDataName)) + " / Value : " + $PubDataValues.Item($PubDataNames.IndexOf($PubDataName)) + " - "
                    }
                    WriteToFile -ExportMode $ExportMode -Add "$([char]2039) Published Data - $PubDataOutput" -NbTab $NbTab
                    }
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
        }
    "6c576f3d-e927-417a-b145-5d3eff9c995f"
    #This is an initialize data activity
        {
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) START ACTIVITY - $ActivityDetailsShort (Activity Type : Initialize Data)" -NbTab $NbTab
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) Parameters were added in the workflow definition" -NbTab $NbTab
        $SqlQuery = "select value, type from CUSTOM_START_PARAMETERS where ParentID = '" + $ActivityID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
                {
                WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) " + $dr["value"] + "=> [" + $dr["type"] + "]$" + $dr["value"].Replace(" ", "_")) -NbTab $NbTab
                }
        $dr.Close()
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
        }
    "fa70125f-267e-4065-a4f6-d5493167d663"
    #This is a return data activity
        {
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) START ACTIVITY - $ActivityDetailsShort (Activity Type : Return Data)" -NbTab $NbTab
        $SqlQuery = "select [Key],Value from PUBLISH_POLICY_DATA where ParentID = '" + $ActivityID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
                {
                $TmpProperty = $dr["value"]
                ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("$([char]2039) " + $dr["key"] + " = ") -TmpProperty $TmpProperty -NbTab $NbTab
                }
        $dr.Close()
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
        $Global:FlagReturnData = $True
        If ($Global:FlagReturnDataList.Contains($ActivityName))
                {$Global:FlagReturnDataNumber[$Global:FlagReturnDataList.IndexOf($ActivityDetailsShort)] = $Global:FlagReturnDataNumber[$Global:FlagReturnDataList.IndexOf($ActivityDetailsShort)] +1}
            else
                {
                $Global:FlagReturnDataList+= $ActivityDetailsShort
                $Global:FlagReturnDataNumber+= 1
                }
        }
    default
    #This is another activity
        {
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) START ACTIVITY - $ActivityDetailsShort (Activity Type : $ActivityTypeName)" -NbTab $NbTab
        $SqlQuery = "select PrimaryDataTable from ObjectTypes, Objects where ObjectTypes.UniqueID = Objects.ObjectType and Objects.UniqueID = '" + $ActivityID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        $TmpTableInit = $False
        while ($dr.Read())
                {
                If ([System.DBNull]::Value.Equals($dr["PrimaryDataTable"]) -eq $False) {
                    $TmpTable = $dr["PrimaryDataTable"]
                    $TmpTableInit = $True
                    }
                }
        $dr.Close()
        If ($TmpTableInit -eq $True)
                {
                $SqlQuery = "SELECT name FROM syscolumns WHERE id = (SELECT id FROM sysobjects WHERE name= '" + $TmpTable + "') ORDER by colorder"
                $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                $dr = $myCommand.ExecuteReader()
                while ($dr.Read())
                    {
                    If ($DoNotExportProperties.Contains($dr["name"]) -eq $False) {
                        $SqlQuery = "select " + $dr["name"] + " from " + $TmpTable + " where uniqueID = '" + $ActivityID + "'"
                        $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                        $dr2 = $myCommand2.ExecuteReader()
                        while ($dr2.Read()) {$TmpProperty = $dr2[$dr["name"]]}
                        $dr2.Close()
                        $TmpPropertyType = "default"
                        If (([System.DBNull]::Value.Equals($tmpProperty) -eq $False) -And ($dr["name"] -eq "Filters")) {$TmpPropertyType = "Filters"}
                        If (([System.DBNull]::Value.Equals($tmpProperty) -eq $False) -And ($TmpProperty.Length -gt 26)) { If ($TmpProperty.Substring(0,26) -eq "<ItemRoot><Entry><FieldId>") {$TmpPropertyType = "Filters"}}
                        If (([System.DBNull]::Value.Equals($tmpProperty) -eq $False) -And ($TmpProperty.Length -gt 29)) { If ($TmpProperty.Substring(0,29) -eq "<ItemRoot><Entry><PropertyId>") {$TmpPropertyType = "QIKProperties"}}
                        switch($TmpPropertyType){
                            "Filters"{
                                $XmlNames = @()
                                $XmlValues = @()
                                $XmlPatterns = @()
                                $XmlRelations = @()
                                $xmlDoc = New-Object System.Xml.XmlDocument
                                $TmpProperty = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" + $TmpProperty
                                [System.Xml.XmlDocument]$xmlDoc.LoadXml($TmpProperty)
                                $Input = New-Object System.Xml.XmlNodeReader $xmlDoc
                                While ($Input.Read()){
                                    If ($Input.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                        switch ($Input.Name){
                                            "FieldID" {$XmlNames+=($Input.ReadString()).Split("/")[1].Split("\")[0]}
                                            "FilterValue" {$XmlValues+=(($Input.ReadString()).Replace("\``~F/", "###DELIM###") -split("###DELIM###"))[1]}
                                            "RelationID"
                                                {
                                                switch(($Input.ReadString()).Split("/")[1].Split("\")[0]){
                                                    "0" {$XmlRelations+="equals"}
                                                    "1" {$XmlRelations+="does not equal"}
                                                    "2" {$XmlRelations+="contains"}
                                                    "3" {$XmlRelations+="does not contain"}
                                                    "4" {$XmlRelations+="matches pattern"}
                                                    "5" {$XmlRelations+="does not match pattern"}
                                                    "6" {$XmlRelations+="less than or equal to"}
                                                    "7" {$XmlRelations+="greater than or equal to"}
                                                    "8" {$XmlRelations+="starts with"}
                                                    "9" {$XmlRelations+="ends with"}
                                                    "10" {$XmlRelations+="less than"}
                                                    "11" {$XmlRelations+="greater than"}
                                                    "13" {$XmlRelations+="after"}
                                                    "14" {$XmlRelations+="before"}
                                                    "default" {$XmlRelations+="unknown filter condition"}
                                                    }
                                                }
                                            }
                                    }
                                }
                                $Input.Close()
                                $XmlOutput = ""
                                ForEach ($XmlName In $XmlNames){
                                    ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("$([char]2039) Filter : $XmlName [" + $XmlRelations.Item($XmlNames.IndexOf($XmlName)) + "]") -TmpProperty $XmlValues.Item($XmlNames.IndexOf($XmlName)) -NbTab $NbTab
                                }
                        }
                            "QIKProperties"{
                                $XmlNames = @()
                                $XmlValues = @()
                                $XmlPatterns = @()
                                $XmlRelations = @()
                                $xmlDoc = New-Object System.Xml.XmlDocument
                                $TmpProperty = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" + $TmpProperty
                                [System.Xml.XmlDocument]$xmlDoc.LoadXml($TmpProperty)
                                $Input = New-Object System.Xml.XmlNodeReader $xmlDoc
                                While ($Input.Read()){
                                    If ($Input.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                        switch ($Input.Name){
                                            "PropertyName" {$XmlNames+=($Input.ReadString()).Split("/")[1].Split("\")[0]}
                                            "PropertyValue" {$XmlValues+=(($Input.ReadString()).Replace("\``~F/", "###DELIM###") -split("###DELIM###"))[1]}
                                            }
                                    }
                                }
                                $Input.Close()
                                $XmlOutput = ""
                                ForEach ($XmlName In $XmlNames){
                                    ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("$([char]2039) $XmlName = ") -TmpProperty $XmlValues.Item($XmlNames.IndexOf($XmlName)) -NbTab $NbTab
                                }
                        }
                        "default"
                        {ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("$([char]2039) " + $dr["name"] + " = ") -TmpProperty $TmpProperty -NbTab $NbTab}
                        }
                        If ($dr["name"]-eq "ScheduleTemplateID") {
                            #This is a Check Schedule activity, let's provide more details about the schedule itself
                            $SqlQuery = "select Name from OBJECTS where uniqueID='{" + $TmpProperty + "}'"
                            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                            $dr2 = $myCommand2.ExecuteReader()
                            while ($dr2.Read()) {WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Schedule name : " + $dr2["Name"]) -NbTab $NbTab}
                            $dr2.Close()
                            $SqlQuery = "select * from SCHEDULES where uniqueID = '{" + $TmpProperty + "}'"
                            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                            $dr2 = $myCommand2.ExecuteReader()
                            while ($dr2.Read())
                                {
                                WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Schedule details : Days of week = " + $dr2["DaysOfWeek"] + " - Days of Month = "+ $dr2["DaysOfMonth"]) -NbTab $NbTab
                                WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Schedule details : Monday = " + $dr2["Monday"] + " - Tuesday = "+ $dr2["Tuesday"] + " - Wednesday = "+ $dr2["Wednesday"] + " - Thursday = "+ $dr2["Thursday"] + " - Friday = "+ $dr2["Friday"] + " - Saturday = "+ $dr2["Saturday"] + " - Sunday = "+ $dr2["Sunday"]) -NbTab $NbTab
                                WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Schedule details : First = " + $dr2["First"] + " - Second = "+ $dr2["Second"] + " - Third = "+ $dr2["Third"] + " - Fourth = "+ $dr2["Fourth"] + " - Last = "+ $dr2["Fourth"]) -NbTab $NbTab
                                WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Schedule details : Days = " + $dr2["Days"] + " - Hours = "+ $dr2["Hours"] + " - Exceptions = "+ $dr2["Exceptions"]) -NbTab $NbTab
                                }
                            $dr2.Close() 
                            $Global:FlagSchedule = $True
                            If ($Global:FlagScheduleList.Contains($ActivityName))
                                {$Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ActivityName)] = $Global:FlagScheduleNumber[$Global:FlagScheduleList.IndexOf($ActivityName)] +1}
                            else
                                {
                                $Global:FlagScheduleList+= $ActivityName
                                $Global:FlagScheduleNumber+= 1
                                }
                            }
                        If ($dr["name"]-eq "CounterID") {
                            #This is a counter activity, let's provide more information on the counter name
                            $SqlQuery = "select Name from OBJECTS where uniqueID='{" + $TmpProperty + "}'"
                            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                            $dr2 = $myCommand2.ExecuteReader()
                            If($dr2.HasRows -ne $false)
                            {
                                while ($dr2.Read()) {
                                    WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Counter name : " + $dr2["Name"]) -NbTab $NbTab
                                }
                            }
                            $dr2.Close()
                            $SqlQuery = "select DefaultValue from COUNTERS where uniqueID = '{" + $TmpProperty + "}'"
                            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                            $dr2 = $myCommand2.ExecuteReader()
                            If($dr2.HasRows -ne $false)
                            {
                                while ($dr2.Read()) {
                                    WriteToFile -ExportMode $ExportMode -Add ("$([char]2039) Counter default value : " + $dr2["DefaultValue"]) -NbTab $NbTab
                                }
                            }
                            $dr2.Close()
                            $Global:FlagCounter = $True
                            If ($Global:FlagCounterList.Contains($ActivityName))
                                {$Global:FlagCounterNumber[$Global:FlagCounterList.IndexOf($ActivityName)] = $Global:FlagCounterNumber[$Global:FlagCounterList.IndexOf($ActivityName)] +1}
                            else
                                {
                                $Global:FlagCounterList+= $ActivityName
                                $Global:FlagCounterNumber+= 1
                                }
                            }
                        If ($dr["name"]-eq "SelectedBranch") {
                            #This is a junction activity, let's stor the activity name to mention in the footer summary
                            $Global:FlagJunction = $True
                            If ($Global:FlagJunctionList.Contains($ActivityName) -eq $False){
                                $Global:FlagJunctionList+= $ActivityName
                                $Global:FlagJunctionNumber+= 1
                                }
                            }
                        }
                    }
                $dr.Close()
                If ($ActivityType -eq "9c1bf9b4-515a-4fd2-a753-87d235d8ba1f"){
                            #This is an invole runbook activity, we also populate the calling parameters for the invoked runbook
                            $SqlQuery = "Select * from TRIGGER_POLICY_PARAMETERS where parentID = '" + $ActivityID + "'"
                            $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
                            $dr = $myCommand.ExecuteReader()
                            while ($dr.Read())
                                    {
                                    $SqlQuery = "select * from CUSTOM_START_PARAMETERS where uniqueID = '" + $dr["parameter"] + "'"
                                    $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection2
                                    $dr2 = $myCommand2.ExecuteReader()
                                    while ($dr2.Read()) {$ParamName = $dr2["value"]}
                                    $dr2.Close()
                                    If ([System.DBNull]::Value.Equals($dr["value"]) -eq $False)
                                        {ParseAndWriteProperty -ExportMode $ExportMode -Prefix ("$([char]2039) Input parameter : " + $ParamName + " = ") -TmpProperty $dr["value"] -NbTab $NbTab}
                                        else {WriteToFile -ExportMode $ExportMode -Add "$([char]2039) Input parameter : $ParamName = < no value was passed >" -NbTab $NbTab}
                                    }
                            $dr.Close()
                            $Global:FlagInvokeRunbook = $True
                            If ($Global:FlagInvokeRunbookList.Contains($ActivityName))
                                {$Global:FlagInvokeRunbookNumber[$Global:FlagInvokeRunbookList.IndexOf($ActivityName)] = $Global:FlagInvokeRunbookNumber[$Global:FlagInvokeRunbookList.IndexOf($ActivityName)] +1}
                            else
                                {
                                $Global:FlagInvokeRunbookList+= $ActivityName
                                $Global:FlagInvokeRunbookNumber+= 1
                                }
                }

        }
        WriteToFile -ExportMode $ExportMode -Add "$([char]2039) END ACTIVITY - $ActivityDetailsShort" -NbTab $NbTab
        }
    }
    $myConnection2.Close()
} #AppendActivityDetails


function ParseRunbookFromActivity() {
# This is the function we recurse on when outputting the structure of the PS1 file,
# as we go through the source Runbook in Orchestrator
# GeneratePS1() calls this function the first time, from the starting activity,
# and then it recurses from there

    param (
    [String]$ActivityID,
    [String]$ActivityName,
    [String]$ActivityType,
    [String]$LinkID,
    [Boolean]$InParallel,
    [Int]$CurrentTabNumber
    )

        $LinkedActivitiesID = @()
        $LinkedActivitiesNames = @()
        $LinkedActivitiesType = @()
        $Links = @()

        $NewTabNumber = $CurrentTabNumber

        $SqlQuery = "select LINKS.TargetObject, Links.UniqueID As LID, OBJECTS.ObjectType, OBJECTS.Name, OBJECTS.UniqueID As OBJID from LINKS, OBJECTS where NOT EXISTS (SELECT UniqueID FROM OBJECTS WHERE UniqueID=LINKS.UniqueID AND DELETED=1) AND LINKS.SourceObject = '" + $ActivityID + "' AND OBJECTS.UniqueID = LINKS.TargetObject AND OBJECTS.Deleted=0"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $LinkedActivitiesID +=$dr["OBJID"]
            $LinkedActivitiesNames +=$dr["Name"]
            $LinkedActivitiesType +=$dr["ObjectType"]
            $Links +=$dr["LID"]
            }
        $dr.Close()
        
        #Let's check if this is the first activity in the runbook
        If ($LinkID -eq "SOURCE")
            {$LinkDetails = ""}
            else {$LinkDetails = LinkCondition($LinkID)}

       #The following is handled differently depending on the number of linked activities 
       switch ($LinkedActivitiesID.Count){
       0
            {
                If ($LinkDetails -ne ""){
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add $LinkDetails -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                AppendActivityDetails -ActivityID $ActivityID -ActivityDetailsShort $ActivityName -ActivityType $ActivityType -NbTab $NewTabNumber -ExportMode "PS1"
                If ($LinkDetails -ne "") {
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
            }
        1
            {
                If (($InParallel -eq $True) -and ($LinkDetails -eq "")) {
                    WriteToFile -ExportMode "PS1" -Add "Sequence {" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                If ($LinkDetails -ne "") {
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add $LinkDetails -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                AppendActivityDetails -ActivityID $ActivityID -ActivityDetailsShort $ActivityName -ActivityType $ActivityType -NbTab $NewTabNumber -ExportMode "PS1"
                ParseRunbookFromActivity -ActivityID $LinkedActivitiesID[0] -ActivityName $LinkedActivitiesNames[0] -ActivityType $LinkedActivitiesType[0] -LinkID $Links[0] -InParallel $False -CurrentTabNumber $NewTabNumber
                If ($LinkDetails -ne "") {
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
                If ($InParallel -eq $True -and $LinkDetails -eq "") {
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
            }
        {$_ -gt 1}
            {
                $Global:FlagParallel = $True
                If ($Global:FlagParallelList.Contains($ActivityName) -eq $False) {$Global:FlagParallelList+= $ActivityName}
                If ($LinkDetails -ne ""){
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add $LinkDetails -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber + 1
                }
                AppendActivityDetails -ActivityID $ActivityID -ActivityDetailsShort $ActivityName -ActivityType $ActivityType -NbTab $NewTabNumber -ExportMode "PS1"
                WriteToFile -ExportMode "PS1" -Add "Parallel {" -NbTab $NewTabNumber
                $NewTabNumber = $NewTabNumber + 1
                $i = 0
                While ($i -lt $LinkedActivitiesID.Count){
                    ParseRunbookFromActivity -ActivityID $LinkedActivitiesID[$i] -ActivityName $LinkedActivitiesNames[$i] -ActivityType $LinkedActivitiesType[$i] -LinkID $Links[$i] -InParallel $True -CurrentTabNumber $NewTabNumber
                    $i = $i + 1
                }
                WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                If ($LinkDetails -ne ""){
                    #Condition was found on link
                    WriteToFile -ExportMode "PS1" -Add "}" -NbTab $NewTabNumber
                    $NewTabNumber = $NewTabNumber - 1
                }
            }
        }
} #ParseRunbookFromActivity


function ListRunbooks() {
# This function is being called when loading the tool
# and everytime the 'Update List' button is being clicked or called
# (the button is also called when hitting enter in the database server
# or database port textboxes)
# The function recurses through the Runbooks in the database
# to fill the TreeView in the GUI by calling the FillNode() function

param (
    [System.Windows.Controls.TreeView]$Tree
)

        $Global:ProgressCount = 0

        Write-Host -ForegroundColor white "["(date -format "HH:mm:ss")"] Trying to connect to server $DatabaseServer"
        Write-Host "Connecting to server $DatabaseServer and retrieving Runbooks..."
        $Tree.Items.Clear()
        $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $eap = $ErrorActionPreference = "SilentlyContinue"
        $myConnection.Open()
        if (!$?)
        {
            $ErrorActionPreference =$eap
            Write-Host "Runbook hierarchy cannot be displayed.`r`nConnection to database server " + $DatabaseServer + " could not be opened.`r`nPlease configure or check the server name on the next screen and try again."
            $Tree.IsEnabled= $False
        }
        else
        {  
            $ErrorActionPreference =$eap
            Write-Host "["(date -format "HH:mm:ss")"] Connected to database, retrieving Runbooks..."
            $NodeRoot = New-Object System.Windows.Controls.TreeViewItem 
            $NodeRoot.Header = "Runbooks"
            $NodeRoot.Name = "Folder"
            $NodeRoot.Tag = "00000000-0000-0000-0000-000000000000"
            [void]$Tree.Items.Add($NodeRoot)
            ## Call Function ##
            FillNode($NodeRoot)

            $Tree.IsEnabled= $True
            Write-Host -ForegroundColor white "["(date -format "HH:mm:ss")"] Runbooks parsing finished..."
            $myConnection.Close()
        }
        Write-Host "Connecting to server $DatabaseServer and retrieving Runbooks..."
} #ListRunbooks


function FillNode() {
# This function is being called by the ListRunbooks() function
# to fill details of a specific folder in the Orchestrator
# hierarchy (subfolders and Runbooks at the root of the folder)
# For subfolders, it actually recurses on itself
# Runbooks are leaf objects in the recursion

param (
    [System.Windows.Controls.TreeViewItem]$TreeNode
)

    Write-Host "Start FillNode Function to retrieve Runbooks..."

    #Retrieve folders
    $SqlQuery = "SELECT Name, UniqueID from FOLDERS WHERE ParentID='" + $TreeNode.Tag + "' AND deleted = 'False' ORDER BY Name"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
    $dr = $myCommand.ExecuteReader()
    while ($dr.Read())
    {
        $NewNode = New-Object System.Windows.Controls.TreeViewItem 
        $NewNode.Header = $dr["Name"]
        $NewNode.Tag = $dr["UniqueID"]
        $NewNode.Name = "Folder"
        [void]$TreeNode.Items.Add($NewNode)
    }
    $dr.Close()

    #Retrieve Runbooks
    $SqlQuery = "select DISTINCT POLICIES.Name AS PName, POLICIES.UniqueID AS PID, FOLDERS.Name As PFName from POLICIES, FOLDERS where FOLDERS.UniqueID = POLICIES.ParentID AND POLICIES.Deleted = 0 AND POLICIES.ParentID = '" + $TreeNode.Tag + "' ORDER BY POLICIES.NAME"
    $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
    $dr = $myCommand.ExecuteReader()
    while ($dr.Read())
    {
        $NewNode = New-Object System.Windows.Controls.TreeViewItem
        $NewNode.Header = $dr["PName"]
        $NewNode.Name = "Runbook"
        $NewNode.Tag = $dr["PID"]
        $NewNode.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Blue)
        [void]$TreeNode.Items.Add($NewNode)
    }
    $dr.Close()

    #Continue recursive search in subfolders
    ForEach ($NewNode In $TreeNode.Items)
    {
        If ($NewNode.Name.Substring(0,6) -eq "Folder")
        {
            FillNode($NewNode)
        }
    }
    Write-Host "End of FillNode Function to retrieve Runbooks..."
} #FillNode


Function Encrypt-Password {
    [CmdletBinding()]
    Param
    ([Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    $Password
    )

    Begin
    {
        $SecurePwd = New-Object -TypeName System.Security.SecureString
    }

    Process
    {
        For($i=0;$i -lt $Password.Length;$i++)
        {
            $SecurePwd.AppendChar($Password[$i])
        }
        $SecurePwd.MakeReadOnly()
        $SecurePwd = ConvertFrom-SecureString $SecurePwd
    }

    End
    {
        Return $SecurePwd
    }
} #Encrypt-Password


Function Decrypt-Password {
    [CmdletBinding()]
    Param
    ([Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    $SecurePassword
    )

    Begin
    {
        $SecurePwd = ConvertTo-SecureString -String $SecurePassword
    }

    Process
    {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePwd)
        $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }

    End
    {
        Return $PlainPassword
    }
} #Decrypt-Password


Function Save-XmlFile {
    [CmdletBinding()]
    Param
    ([Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    $XmlFilePath,
    [Parameter(Mandatory=$true)]
    $XmlDocument
    )

    If(Test-Path -Path:$ConfigurationPath -IsValid)
    {
        Try {
            ## Save file using a textwriter and indented formating ##
            $TextWriter = New-Object -TypeName System.Xml.XmlTextWriter($XmlFilePath,[System.Text.Encoding]::UTF8)
            $TextWriter.Formatting = [System.Xml.Formatting]::Indented
            $XmlDocument.Save($TextWriter)
            $TextWriter.Dispose()

            write-host -ForegroundColor Gray "["(date -format "HH:mm:ss")"] Successful saving XML File modification for : $XmlFilePath"
        }
        Catch {
            write-host -ForegroundColor Red "["(date -format "HH:mm:ss")"] Error on saving XML File at : $XmlFilePath"
            write-host -ForegroundColor Red "["(date -format "HH:mm:ss")"] Error message: $($Error[0])"
        }
    }
} #Save-XmlFile


Function Generate-SingleRunbookName {
Param([String]$RunbookName, [Int]$Depth, $RunbookId)

    $SimplifiedRunbookName = $RunbookName.Replace(" ", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("/", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("\", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace(">", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("<", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace(":", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("*", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("?", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("|", "")
    $SimplifiedRunbookName = $SimplifiedRunbookName.Replace("-", "")

    ## Add the Depth at the End ##
    $SimplifiedRunbookName = "$SimplifiedRunbookName" + "-$Depth" + "-$RunbookId" + "-$(Generate-BasicDate)"

    Return $SimplifiedRunbookName
} #Generate-SingleRunbookName


Function Read-TreeView {
param($Tree, [Int]$Depth, $VisioApp, [Boolean]$VisioState)
    
    $Depth = $Depth + 1

    If($Tree.HasItems)
    {
        ## Get the list of all Items which are Runbook ##
        $MyRunbookList = $Tree.Items | Where-Object {$_.Name -eq "Runbook"} | Sort-Object -Property Name
        $MyFolderList = $Tree.Items | Where-Object {$_.Name -eq "Folder"} | Sort-Object -Property Name
        
        If($MyRunbookList.Count -gt 0)
        {
            ForEach($RunbookItem in $MyRunbookList)
            {
                If(($Global:VisioAvailability) -and ($Depth -gt 1))
                {
                    $SimplifiedRunbookName = Generate-SingleRunbookName -RunbookName $RunbookItem.Header -RunbookId $RunbookItem.Tag  -Depth $Depth
                    $VisioFilePath = Generate-Visio -RunbookID $RunbookItem.Tag -RunbookName $RunbookItem.Header -ExportFileName $SimplifiedRunbookName -vApp:$VisioApp
                }
                Generate-WordDoc -RunbookID:$RunbookItem.Tag -RunbookName:$RunbookItem.Header -ObjectType:$RunbookItem.Name -oDoc:$WordDoc -Depth:$Depth -VisioPath:$VisioFilePath
            }
        }

        If($MyFolderList.Count -gt 0)
        {
            ForEach($FolderItem in $MyFolderList)
            {
                If($FolderItem.HasItems)
                {
                    Generate-WordDoc -RunbookID:$FolderItem.Tag -RunbookName:$FolderItem.Header -ObjectType:$FolderItem.Name -oDoc:$WordDoc -Depth:$Depth
                    Read-TreeView -Tree:$FolderItem -Depth:$Depth -VisioApp:$VisioApp -VisioState:$VisioState
                }
                Else
                {
                    Generate-WordDoc -RunbookID:$FolderItem.Tag -RunbookName:$FolderItem.Header -ObjectType:$FolderItem.Name -oDoc:$WordDoc -Depth:$Depth
                }
            }
        }
    }
} #Read-TreeView


Function Generate-WordDoc {
param (
    [String]$RunbookID,
    [String]$RunbookName,
    [String]$ObjectType,
    [String]$VisioPath,
    [Int]$Depth,
    $oDoc
)

    $Global:FlagParallel = $False
    $Global:FlagParallelList = @()
    $Global:FlagJunction = $False
    $Global:FlagJunctionList = @()
    $Global:FlagStringcondition = $False
    $Global:FlagInitializeData = $False
    $Global:FlagVariables = $False
    $Global:FlagVariablesList = @()
    $Global:FlagVariablesNumber = @()
    $Global:FlagVariablesValue =@()
    $Global:FlagInvokeRunbook = $False
    $Global:FlagInvokeRunbookList = @()
    $Global:FlagInvokeRunbookNumber = @()
    $Global:FlagReturnData = $False
    $Global:FlagReturnDataList = @()
    $Global:FlagReturnDataNumber = @()
    $Global:FlagSchedule = $False
    $Global:FlagScheduleList = @()
    $Global:FlagScheduleNumber = @()
    $Global:FlagCounter = $False
    $Global:FlagCounterList = @()
    $Global:FlagCounterNumber = @()

    $myConnectionDOC = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
    $myConnectionDOC.Open()
    $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
    $myConnection.Open()
    $myConnection2 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
    $myConnection2.Open()

    #Start Word and open the document template.
    #Insert a paragraph at the beginning of the document.
    $oPara1 = $oDoc.Content.Paragraphs.Add()

    ## Word contains only 9 different style for Titles, if 9 is reached put 999 as Value, apply Title 9 and modify the Name of the folder with (Style Not applied)
    If($ObjectType -eq "Folder")
    {
        If($Depth -le 8)
        {
            $oPara1.Range.Text = "$ObjectType : $RunbookName"
            $oPara1.Style = $Global:TitleStyleArray | Where-Object {$_.NameLocal -match "$Depth$"}
        }
        Else
        {
            $oPara1.Range.Text = "$ObjectType : $RunbookName (Folder too deep, Style not applied, depth: $Depth)"
            $oPara1.Style = $Global:MaxTitleDepth
        }

        $oPara1.Range.Font.Bold = $True
        $oPara1.SpaceBefore = 24
        $oPara1.SpaceAfter = 12
        $oPara1.Range.InsertParagraphAfter()
    }
    Else
    {
        $oPara1.Range.Text = "$ObjectType : $RunbookName"
        $oPara1.Style = $Global:RunbookWordStyle
        $oPara1.Range.Font.Bold = $True
        $oPara1.SpaceBefore = 5
        $oPara1.SpaceAfter = 5
    }

    If($ObjectType -ne "Folder")
    {
        If(-not ([String]::IsNullOrWhiteSpace($VisioPath)))
        {
            ## Put Visio inside the Doc ##
            $oPara1 = $oDoc.Content.Paragraphs.Add()
            $oPara1.SpaceBefore = 5
            $oPara1.SpaceAfter = 5
            $oTable = $oDoc.Tables.Add($oDoc.Bookmarks.Item("\endofdoc").Range, 1, 2)
            $oTable.Range.Font.Size = 8
            $oTable.Range.Font.Bold = $True
            $oTable.Range.Borders.Enable = $False
            $oTable.Cell(1, 1).Range.Text = "Runbook Visio File Name : "
            ## Insert Visio name Now, keep the path for futur improvment ##
            $VisioName = [System.IO.Path]::GetFileName($VisioPath)
            $oTable.Cell(1, 2).Range.Text = "$VisioName"

            $oTable.Columns.Item(1).Width = $oDoc.Parent.InchesToPoints(1.5)
            $oTable.Columns.Item(2).Width =$oDoc.Parent.InchesToPoints(6)
        }

        $oPara1 = $oDoc.Content.Paragraphs.Add()
        $oPara1.LineSpacing = 1
        $oPara1.SpaceBefore = 0
        $oPara1.SpaceAfter = 5
        #Insert a 1 x 4 table, fill it with data, and make the first row bold and italic.
        $oTable = $oDoc.Tables.Add($oDoc.Bookmarks.Item("\endofdoc").Range, 1, 4)
        $oTable.Range.ParagraphFormat.SpaceAfter = 12
        $oTable.Range.Font.Size = 8
        $oTable.Range.Font.Bold = $True
        $oTable.Range.Borders.Enable = $True
        $oTable.Range.Borders.OutsideLineStyle = 1
        $oTable.Range.Borders.InsideLineStyle = 0
        $oTable.Columns.Borders.InsideLineStyle = 1
        $oTable.Cell(1, 1).Range.Text = "Activity"
        $oTable.Cell(1, 2).Range.Text = "Description"
        $oTable.Cell(1, 3).Range.Text = "Details"
        $oTable.Cell(1, 4).Range.Text = "Published data dependencies"
        $oTable.Columns.Item(1).Width = $oDoc.Parent.InchesToPoints(1.5)
        $oTable.Columns.Item(3).Width = $oDoc.Parent.InchesToPoints(4)
        $oTable.Columns.Item(2).Width = $oDoc.Parent.InchesToPoints(1)
        $oTable.Columns.Item(4).Width = $oDoc.Parent.InchesToPoints(1)

        $r = 2
        ## Add an order by CASE, "ORDER BY OBJECTS.PositionY ASC, OBJECTS.PositionX ASC" same as the representation in Orchestrator ##
        $SqlQueryDOC = "select UniqueID, ObjectType, Name, Description, PositionX, PositionY from OBJECTS WHERE ParentID = '" + $RunbookID + "' AND ObjectType <> '7A65BD17-9532-4D07-A6DA-E0F89FA0203E' AND Deleted=0 ORDER BY OBJECTS.PositionY ASC, OBJECTS.PositionX ASC"
        $myCommandDOC = New-Object System.Data.SqlClient.sqlCommand $SqlQueryDOC, $myConnectionDOC
        $drDOC = $myCommandDOC.ExecuteReader()
        while ($drDOC.Read())
            {
            [void]$oTable.Rows.Add()
            $oTable.Rows.Item($r).Range.Font.Bold = $False
            $oTable.Rows.Item($r).Range.Borders.Enable = $True
            $oTable.Rows.Item($r).Range.Borders.OutsideLineStyle = 1
            $oTable.Rows.Item($r).Range.Borders.InsideLineStyle = 0
            $oTable.Cell($r, 1).Range.Text = $drDOC["Name"]
            $oTable.Cell($r, 1).Range.Paragraphs.SpaceAfter = 0
            $oPara1 = $oTable.Cell($r, 1).Range.Paragraphs.Add()
            $oPara1.SpaceAfter = 0
            If (Test-Path ("$ActivitiesPicturePath" +"{" + $drDOC["ObjectType"] + "}.jpg"))
                {[void]$oPara1.Range.InlineShapes.AddPicture(("$ActivitiesPicturePath" + "\{" + $drDOC["ObjectType"] + "}.jpg"))}
                else {[void]$oPara1.Range.InlineShapes.AddPicture(("$ActivitiesPicturePath" + "\default.jpg"))}

            $oPara1.SpaceAfter = 0

            If ([System.DBNull]::Value.Equals($drDOC["description"]) -eq $False)
                {
                $oTable.Cell($r, 2).Range.Text = $drDOC["Description"]
                $oTable.Cell($r, 2).Range.Paragraphs.SpaceAfter = 0
                }

            $SqlQuery2 = "select PrimaryDataTable from ObjectTypes, Objects where ObjectTypes.UniqueID = Objects.ObjectType and Objects.UniqueID = '" + $drDOC["UniqueID"] + "'"
            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery2, $myConnection2
            $dr2 = $myCommand2.ExecuteReader()
            $TmpTableInit = $False
            while ($dr2.Read())
                {
                If ([System.DBNull]::Value.Equals($dr2["PrimaryDataTable"]) -eq $False)
                    {
                    $TmpTable = $dr2["PrimaryDataTable"]
                    $TmpTableInit = $True
                    }
                }
            $dr2.Close()
            $Global:ActivityDependenciesActivityNames = @()
            $Global:ActivityDependenciesActivityTypes = @()
            $Global:ActivityDependenciesVariableNames = @()
            $Global:ActivityDependenciesVariableValues = @()
            AppendActivityDetails -ActivityID ($drDOC["UniqueID"]) -ActivityDetailsShort ($drDOC["Name"]) -ActivityType ($drDOC["ObjectType"]) -ExportMode "DOC"

            ## Create a Boolean to remove blank Space for first value ##
            [Boolean]$FirstActivityName = $True

            foreach ($ActivityDependencyActivityName in $Global:ActivityDependenciesActivityNames)
                {
                If($FirstActivityName)
                {
                    $oTable.Cell($r, 4).Range.Text = $ActivityDependencyActivityName
                    $oTable.Cell($r, 4).Range.Paragraphs.SpaceAfter = 0
                }
                Else
                {
                    $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                    $oPara1.Range.Text = $ActivityDependencyActivityName
                    $oPara1.SpaceAfter = 0
                }

                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                $oPara1.SpaceAfter = 0
                If (Test-Path ("$ActivitiesPicturePath{" + $Global:ActivityDependenciesActivityTypes[$Global:ActivityDependenciesActivityNames.IndexOf($ActivityDependencyActivityName)] + "}.jpg"))
                        {[void]$oPara1.Range.InlineShapes.AddPicture(("$ActivitiesPicturePath"+ "{" + $Global:ActivityDependenciesActivityTypes[$Global:ActivityDependenciesActivityNames.IndexOf($ActivityDependencyActivityName)] + "}.jpg"))}
                        else {[void]$oPara1.Range.InlineShapes.AddPicture(("$ActivitiesPicturePath" + "default.jpg"))}
                }
            foreach ($ActivityDependenciesVariableName in $Global:ActivityDependenciesVariableNames)
                {
                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                $oPara1.Range.Text = (" Variable: " + $ActivityDependenciesVariableName)
                $oPara1.SpaceAfter = 0
                $oPara1 = $oTable.Cell($r, 4).Range.Paragraphs.Add()
                $oPara1.Range.Text = (" Value: " + $Global:ActivityDependenciesVariableValues[$Global:ActivityDependenciesVariableNames.IndexOf($ActivityDependenciesVariableName)])
                $oPara1.SpaceAfter = 0
                }
            $r = $r + 1
            }

        $drDOC.Close()
    }

    $myConnectionDOC.Close()
    $myConnection.Close()
    $myConnection2.Close()
    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting Runbook '$RunbookName' to Word file"

} #Generate-WordDoc


Function Generate-BasicDate {

    $Today = (Get-Date)
    $Year = $Today.Year.ToString()
    $Month = $Today.Month.ToString()
    $Day = $Today.Day.ToString()
    $Hour = $Today.Hour.ToString()
    $Minute = $Today.Minute.ToString()

    If($Month.Length -lt 2)
    {
        $Month = "0"+$Month
    }

    If($Day.Length -lt 2)
    {
        $Day = "0"+$Day
    }

    If($Hour.Length -lt 2)
    {
        $Hour = "0"+$Hour
    }

    If($Minute.Length -lt 2)
    {
        $Minute = "0"+$Minute
    }

    $BasicDate = "$Year-$Month-$Day-$Hour-$Minute"

    Return $BasicDate
} #Generate-BasicDate


Function Generate-Visio {
    param (
        [String]$RunbookID,
        [String]$RunbookName,
        [String]$ExportFileName,
        $vApp
    )

        $myConnection = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection.Open()
        $myConnection2 = New-Object System.Data.SqlClient.SqlConnection $Global:SQLConnstr 
        $myConnection2.Open()

        $ListShapes = @()
        $ListShapes.Clear()
        $ListShapesID = @()
        $ListShapesID.Clear()

        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Creating header"
        $vDoc = $vApp.Documents.Add("")
        $vStencil = $vApp.Documents.OpenEx($VisioTemplate, 4)
        $SpecificvStencil = $vStencil.Masters | where-object {$_.NameU -eq "Process"}
        
        ## Get the Page ##
        $VisioPages = $vApp.ActiveDocument.Pages
        $VisioPage = $VisioPages.Item(1)

        #Size the page dynamically
        $VisioPage.AutoSize = $true

        #Add a title
        $vToShape = $VisioPage.Drop($SpecificvStencil, 4, 10)
        $vToShape.Text = "Runbook : $RunbookName"
        $vToShape.Cells("Char.Size").Formula = "= 30 pt."
        $vToShape.Cells("Width").Formula = "= 7"

        #Find and draw activities
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing activities"
        $vFlowChartMaster = $vStencil.Masters | where-object {$_.NameU -eq $Global:VisioStencil}

        $SqlQuery = "select UniqueID, ObjectType, Name, Description, PositionX, PositionY from OBJECTS WHERE ParentID = '" + $RunbookID + "' AND ObjectType <> '7A65BD17-9532-4D07-A6DA-E0F89FA0203E' AND Deleted=0"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Working with activity :" $dr["Name"]
            $vApp.ActiveWindow.DeselectAll()
            $vToShape = $VisioPage.Drop($vFlowChartMaster, ($dr["PositionX"] / 50), (-($dr["PositionY"] / 50) + 7))
            $vToShape.Text = $dr["Name"]
            If ([System.DBNull]::Value.Equals($dr["description"]) -eq $False)
                {
                $vsoDoc1 = $vApp.Documents.OpenEx($vApp.GetBuiltInStencilFile(3, 0), [Microsoft.Office.Interop.Visio.VisMeasurementSystem]::visMSDefault)
                $vCallout = $VisioPage.DropCallout($vsoDoc1.Masters.ItemU($VisioCallout), $vToShape)
                $vCallout.Text = $dr["Description"]
                $vsoDoc1.Close()
                }
            $vToShape.Cells("Para.HorzAlign").Formula = "=2"
            $vToShape.Cells("LeftMargin").Formula = "=0.5"

            #get the icon for this activity
            If (Test-Path ("$ActivitiesPicturePath{" + $dr["ObjectType"] + "}.jpg"))
                {$shp1Obj = $VisioPage.Import("$ActivitiesPicturePath" + "{" + $dr["ObjectType"] + "}.jpg")}
                else {$shp1Obj = $VisioPage.Import("$ActivitiesPicturePath" + "\default.jpg")}
            $shp2Obj = $VisioPage.Drop($shp1Obj, ($dr["PositionX"] / 50) - 0.25, -($dr["PositionY"] / 50) + 7)
            $shp1Obj.Delete() #Remove original imported reference
            $vApp.ActiveWindow.Select($vToShape, 2)
            $vApp.ActiveWindow.Select($shp2Obj, 2)
            $vSel = $vApp.ActiveWindow.Selection
            [void]$vSel.Group()
            $ListShapes += $vToShape
            $ListShapesID += $dr["UniqueID"]
            $vSel.DeselectAll()
            }
        $dr.Close()

        #Find and draw links
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing links"
        $vConnectorMaster = $vStencil.Masters | where-object {$_.NameU -eq "Dynamic Connector"}
        $SqlQuery = "Select DISTINCT LINKS.UniqueID As LID, name, deleted, objecttype, LINKS.Color, LINKS.sourceobject, LINKS.targetobject from OBJECTS, LINKS where (ObjectType='7A65BD17-9532-4D07-A6DA-E0F89FA0203E' AND ParentID='" + $RunbookID + "' AND OBJECTS.Deleted=0 AND LINKS.UniqueID=OBJECTs.UniqueID)"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $vConnector = $VisioPage.Drop($vConnectorMaster, 0, 0)
            $vConnector.Cells("EndArrow").Formula = "=4"

            ## Color part ##
            $TmpColor=[Long]$dr["Color"]
            $TmpRed = $TmpColor % 256
            $TmpColor = $TmpColor / 256
            $TmpGreen = $TmpColor % 256
            $TmpColor = $TmpColor / 256
            $TmpBlue = $TmpColor % 256

            Try {$vConnector.Cells("LineColor").Formula = "RGB($TmpRed;$TmpGreen;$TmpBlue)"}
            Catch {
                $vConnector.Cells("LineColor").Formula = "RGB($TmpRed,$TmpGreen,$TmpBlue)"
            }
            $vBeginCell = $vConnector.Cells("BeginX")
            $vFromShape = $ListShapes.Item($ListShapesID.IndexOf($dr["SourceObject"]))
            $vBeginCell.GlueTo($vFromShape.Cells("Align" + $Global:VisioGlueFrom))
            $vEndCell = $vConnector.Cells("EndX")
            $vToShape = $ListShapes.Item($ListShapesID.IndexOf($dr["TargetObject"]))
            $vEndCell.GlueTo($vToShape.Cells("Align" + $Global:VisioGlueTo))
            #LID to String?
            $SqlQuery2 = "select DISTINCT Name, LINKS.UniqueID from LINKS, OBJECTS WHERE OBJECTS.UniqueID = LINKS.UniqueID AND LINKS.UniqueID = '" + $dr["LID"] + "'"
            $myCommand2 = New-Object System.Data.SqlClient.sqlCommand $SqlQuery2, $myConnection2
            $dr2 = $myCommand2.ExecuteReader()
            while ($dr2.Read())
                {
                If ($dr2["name"] -ne "Link") {$vConnector.Text = $dr2["Name"]}
                }
            $dr2.Close()
            #$myCommand2 = Nothing
            $vConnector.SendToBack()
            }
        $dr.Close()

        #Find and draw loops
        write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Drawing loops"
        $SqlQuery = "select Objectlooping.UniqueID As OUID, DelaybetweenAttempts, Name from objectlooping, objects where objects.uniqueid = objectlooping.uniqueID and objectlooping.enabled = 1 AND OBJECTS.Deleted=0 AND objects.parentID = '" + $RunbookID + "'"
        $myCommand = New-Object System.Data.SqlClient.sqlCommand $SqlQuery, $myConnection
        $dr = $myCommand.ExecuteReader()
        while ($dr.Read())
            {
            $vConnector = $VisioPage.Drop($vConnectorMaster, 0, 0)
            $vConnector.Cells("EndArrow").Formula = "=4"
            $vBeginCell = $vConnector.Cells("BeginX")
            $vFromShape = $ListShapes.Item($ListShapesID.IndexOf($dr["OUID"]))
            $vBeginCell.GlueTo($vFromShape.Cells("AlignRight"))
            $vEndCell = $vConnector.Cells("EndX")
            $vToShape = $ListShapes.Item($ListShapesID.IndexOf($dr["OUID"]))
            $vEndCell.GlueTo($vToShape.Cells("AlignTop"))
            Try {If ($dr["DelayBetweenAttempts"] -ne "") {$vConnector.Text = "Loop every " + $dr["DelayBetweenAttempts"] + " seconds"}}
            Catch {$vConnector.Text = "Loop (undefined interval)"}
            $vConnector.SendToBack()
            }
        $dr.Close()

        $VisioPage.AutoSizeDrawing()

        $myConnection.Close()
        $myConnection2.Close()

        ## Generate unique file Path ##
        $ExportFilePath = "$OutputPath" + "$ExportFileName" + ".VSDX"

    write-host  -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting Runbook '$RunbookName' to Visio file at path: $ExportFilePath"
    [void]$vDoc.SaveAs($ExportFilePath)
    $vDoc.Close()

    Return $ExportFilePath

} #Generate-Visio


#################################################################################
#                                                                               #
#                               Launch Main Code                                #
#                                                                               #
#################################################################################

## Move to the Script Path ##
Set-Location $pathScript

## Thanks to http://gallery.technet.microsoft.com/scriptcenter/63fd1c0d-da57-4fb4-9645-ea52fc4f1dfb ##
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
If(-not $IsAdmin)
{
    Try
    {
        $RelaunchArgument = "-file `"$($ScriptFilePath)`"" 
        Write-Host -ForegroundColor Yellow "["(date -format "HH:mm:ss")"] WARNING : This script should run with administrative rights - Relaunching the script in elevated mode in 3 seconds..."
        Start-Sleep 3
        Start-Process "$PSHOME\powershell.exe" -Verb Runas -ArgumentList $RelaunchArgument -ErrorAction Stop
    }
    Catch
    {
        write-host -ForegroundColor Red "["(date -format "HH:mm:ss")"] Error : Failed to restart script with administrative rights - please make sure this script is launched elevated."
        Break
    }
    Exit
}
Else
{
    write-host -ForegroundColor Gray "["(date -format "HH:mm:ss")"] We are running in elevated mode, we can proceed with launching the tool."
}

## Try to create a Word Document, if failed Word is certainly not installed ##
Try {
    ## Create a new document ##
    $WordFile = New-Object -ComObject Word.Application

    #Start Word and open the document template.
    $WordFile.Visible = $True
    $WordDoc = $WordFile.Documents.Add()
    Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] Creation of the Word Document"
    
    ## We manage Style for Only English and French languages, else we fail until someone update the script :) ##
    Switch ($WordFile.Language.value__)
    {
        ## For New language, just copy paste the next 4 line, indend the number with the word language value and modify the word "title" with the right word in your language ##
        1036    {
                    ## C'est du français ##
                    $TitleRegexp = 'Titre [1-8]$'
                    $RunbookStyleRegexp = 'Titre 9'
                }
        1033    {
                    ## This is english ##
                    $TitleRegexp = 'Title [1-8]$'
                    $RunbookStyleRegexp = 'Title 9'
                }
    }
    ## Get Word Style ##
    $Global:TitleStyleArray = $WordDoc.Styles._NewEnum | where {$_.NameLocal -match "$TitleRegexp"} | Sort-Object -Property NameLocal
    $Global:MaxTitleDepth = $TitleStyleArray[($TitleStyleArray.Count - 1)]
    $Global:RunbookWordStyle = $WordDoc.Styles._NewEnum | where {$_.NameLocal -match "$RunbookStyleRegexp"} | Sort-Object -Property NameLocal

    ## Put low margin to have a bigger array ##
    $WordDoc.PageSetup.TopMargin = 25
    $WordDoc.PageSetup.RightMargin = 25
    $WordDoc.PageSetup.LeftMargin = 25
    $WordDoc.PageSetup.BottomMargin = 25
}
Catch {
    Write-Host -ForegroundColor Red "["(date -format "HH:mm:ss")"] An error occured during instanciation of a new Word document, please insure that Word Application is installed"
    Exit
}

## We will try to create a visio to check if the component is installed ##
Try {
    ## Create a new document ##
    $Global:VisioApp = New-Object -ComObject Visio.InvisibleApp
    $Global:VisioAvailability = $True
    $Global:VisioApp.AlertResponse = 7
    Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] We will generate for each Runbook a dedicated Visio file"
}
Catch {
    ## Visio not installed ##
    $Global:VisioAvailability = $False
    Write-Host -ForegroundColor Orange "["(date -format "HH:mm:ss")"] Visio has not been detected, it will not be possible to generate Visio files"
    Continue
}

## Check If Secure Password ##
If([String]::IsNullOrWhiteSpace($DatabaseSecurePassword) -and (-not [String]::IsNullOrWhiteSpace($DatabasePassword)))
{
    ## Need to Encrypt Clear Password ##
    $DatabaseSecurePassword = Encrypt-Password -Password:$DatabasePassword
    ## Update the SecurePassword in the configuration XML ##
    $XmlConfiguration.DefaultConfiguration.DatabaseSecurePassword = "$DatabaseSecurePassword"
    ## Remove the Clear Password from the configuration XML ##
    $XmlConfiguration.DefaultConfiguration.DatabasePassword = ""
}
ElseIf((-not [String]::IsNullOrWhiteSpace($DatabaseSecurePassword)) -and (-not [String]::IsNullOrWhiteSpace($DatabasePassword)))
{
    Try {
        $DecryptedPwd = Decrypt-Password -SecurePassword:$DatabaseSecurePassword
    }
    Catch {
        $DecryptedPwd = ''
        Continue
    }

    ## Decrypt the Secure Password and try to match with the Clear password ##
    If($DatabasePassword -eq $DecryptedPwd)
    {
        ## Just remove the Clear Password ##
        $XmlConfiguration.DefaultConfiguration.DatabasePassword = ""
    }
    Else
    {
        ## Update the SecurePassword in the configuration XML ##
        $XmlConfiguration.DefaultConfiguration.DatabaseSecurePassword = "$(Encrypt-Password -Password:$DatabasePassword)"
        ## Remove the Clear Password from the configuration XML ##
        $XmlConfiguration.DefaultConfiguration.DatabasePassword = ""
    }
}
ElseIf((-not [String]::IsNullOrWhiteSpace($DatabaseSecurePassword)) -and ([String]::IsNullOrWhiteSpace($DatabasePassword)))
{
    ## Decrypt Secure Password ##
    $DatabasePassword = $(Decrypt-Password -SecurePassword:$DatabaseSecurePassword)
}
Else
{
    Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] No password detected, "
}

## Save the configuration ##
Save-XmlFile -XmlFilePath:$($MyConfigurationXmLFile.FullName) -XmlDocument:$XmlConfiguration
###################################

## Build the Sql Connection String ##
If((-not [String]::IsNullOrWhiteSpace($DatabaseServer)) -or (-not [String]::IsNullOrWhiteSpace($DatabasePort)) -or (-not [String]::IsNullOrWhiteSpace($DatabaseName)))
{
    If(-not [String]::IsNullOrWhiteSpace($DatabaseUserName))
    {
        ## We will use a set of credential for Authentication ##
        $Global:SQLConnstr = "Server=$DatabaseServer,$DatabasePort;User Id=$DatabaseUserName;Password=$DatabasePassword;Database=$DatabaseName"
        Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] Script will use SQL Server authentication"
    }
    Else
    {
        $Global:SQLConnstr = "Server=" + $DatabaseServer +"," +  $DatabasePort + ";Integrated Security=SSPI;database=" + $DatabaseName
        Write-Host -ForegroundColor Green "["(date -format "HH:mm:ss")"] Script will use Kerberos authentication"
    }
}
Else
{
    Write-Host -ForegroundColor Red "["(date -format "HH:mm:ss")"] Please control the Database confiugration in the configuration File"
    Exit
}
###################################

##########################################################################################
# Current start of the main process
##########################################################################################

Write-Host -ForegroundColor green "["(date -format "HH:mm:ss")"] Orchestrator Design Dossier Creator, version: $ToolVersion"

## Create a Treeview Object ##
$Tree = New-Object System.Windows.Controls.TreeView

## Get the List of all know Folder and Runbook ##
Write-Host -ForegroundColor white "["(date -format "HH:mm:ss")"] Scan Orchestrator Database to create Runbook tree"
ListRunbooks($Tree)

## Iteration over all Tree Nodes, to create a Doc file ##
Write-Host -ForegroundColor white "["(date -format "HH:mm:ss")"] Start Read Treeview"
Read-TreeView -Tree:$Tree -Depth 0 -VisioApp:$VisioApp -VisioState:$VisioAvailability -WordDocument:$WordDoc

## Save word document and quit apps ##
Write-Host -ForegroundColor white "["(date -format "HH:mm:ss")"] Saving Word Doc file in the Output Folder"
$WordDoc.SaveAs([ref]("$OutputPath" + "Orchestrator-Design-Dossier-" + $(Generate-BasicDate) + ".DOCX"))
$WordDoc.Close()
$WordFile.Quit()
$Global:VisioApp.Quit()

##########################################################################################