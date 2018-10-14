###############################################################################
#                                                                             #
#                        BI Publisher Archive Manager                         #
#                                                                             #
#                                                                             #
# Author  : Francois CABANNES                                                 #
# Version : 1.0.0                                                             # 
# Date    : 14/10/2018                                                        #
#                                                                             #
###############################################################################
#                                                                             #
###############################################################################

# Test : e:\OneDrive\Projects\BIPVersioner\source\extractor.ps1 -InputPath "C:/Temp/AA/in" -OutputPath "C:/Temp/AA/out" -WorkPath "C:/Temp/AA/tmp"

#
# Parse Parameters
#   1 - Input Folder Path
#   2 - Output Folder Path
#   3 - Temporary Folder Path
#
[CmdletBinding()]
Param(
    # Optional Configuration File for default values
    [Parameter(Mandatory = $False)]
    [Alias("Config")]
    [string]$pConfig,

    # Input folder for files to process
    [Parameter(Mandatory = $False)]
    [Alias("InputPath")]
    [string]$pInputDirPath,

    # Output folder for all the processed and resulting files
    [Parameter(Mandatory = $False)]
    [Alias("OutputPath")]
    [string]$pOutputDirPath,

    # Temporary folder used for the processing.
    # Recursively purge before processing !
    [Parameter(Mandatory = $False)]
    [Alias("WorkPath")]
    [string]$pWorkDirPath,

    # Switch to bypass the extraction and processing of security related files
    [Parameter(Mandatory = $False)]
    [Alias("OmitSecurity")]
    [switch]$pOmitSecurity,

    # Switch to bypass the extraction and processing of report thumbnail files
    [Parameter(Mandatory = $False)]
    [Alias("OmitThumbnails")]
    [switch]$pOmitThumbnails,

    # Switch to bypass the extraction and processing of all metadata files
    [Parameter(Mandatory = $False)]
    [Alias("OmitMetadata")]
    [switch]$pOmitMetadata,

    # Switch to bypass the extraction and processing of all metadata files
    [Parameter(Mandatory = $False)]
    [Alias("OmitDataSamples")]
    [switch]$pOmitDataSamples,     

    # Switch to bypass the extraction and processing of all Report folders
    # This means that all Report related files will be ignored.
    [Parameter(Mandatory = $False)]
    [Alias("OmitReports")]
    [switch]$pOmitReports,     

    # Switch to bypass the extraction and processing of all Data Model folders
    # This means that all Data Model related files will be ignored.
    [Parameter(Mandatory = $False)]
    [Alias("OmitDataModels")]
    [switch]$pOmitDataModels,     

    # Switch to bypass the extraction and processing of all Sub Templates files
    [Parameter(Mandatory = $False)]
    [Alias("OmitSubtemplates")]
    [switch]$pOmitSubtemplates,     

    # Switch to bypass the extraction and processing of all Translation files
    [Parameter(Mandatory = $False)]
    [Alias("OmitTranslations")]
    [switch]$pOmitTranslations,   

    # Switch to bypass the extraction and processing of all Style Templates files
    [Parameter(Mandatory = $False)]
    [Alias("OmitStyleTemplates")]
    [switch]$pOmitStyleTemplates,

    # Switch to enable the extraction as single files of all SQL statements in DataModels
    [Parameter(Mandatory = $False)]
    [Alias("ExtractSQL")]
    [switch]$pExtractSQL   

)

# Constants
Set-Variable ZIP_CMDPATH -value "C:/Program Files/7-Zip/7z.exe"

#
#
#
function loadConfiguration ($pConfigFile) {
    # Import email settings from config file
    [xml]$ConfigFile = Get-Content ($pConfigFile)
    $Settings = @{
        ZipSettings = @{
            Path = $ConfigFile.Settings.ZipSettings.Path
        }
        AppSettings = @{
            Certs = $ConfigFile.Settings.AppSettings.Certs
        }
        Arguments   = @{
            InputDirPath       = $ConfigFile.Settings.Arguments.InputDirPath
            OutputDirPath      = $ConfigFile.Settings.Arguments.OutputDirPath
            WorkDirPath        = $ConfigFile.Settings.Arguments.WorkDirPath
            OmitThumbnails     = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitThumbnails)
            OmitDataSamples    = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitDataSamples)
            OmitSecurity       = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitSecurity)
            OmitMetadata       = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitMetadata)
            OmitReports        = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitReports)
            OmitDataModels     = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitDataModels)
            OmitSubtemplates   = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitSubtemplates)
            OmitTranslations   = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitTranslations)
            OmitStyleTemplates = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.OmitStyleTemplates)
            ExtractSQL         = [System.Convert]::ToBoolean($ConfigFile.Settings.Arguments.ExtractSQL)
        }
    }
    return $Settings  
}


#
# Displays a Yes/No command-line prompt with a custom message
#
function PromptYesNo ($msg) {
    do {
        $response = Read-Host -Prompt $msg
        if ($response -eq 'y') {        
            return $true
        }
    } until ($response -eq 'n')
    return $false
}

#
# Check for a folder existance. If id does not exist prompt for creation.
# Returns true is the folder already exists or was created succesfully. Else False.
#
function CheckFolderAndCreate {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [string]$path,
        [Parameter(Mandatory = $True)]
        [string]$msg
    )

    if (!(Test-Path -PathType Any "$path")) {
        if (PromptYesNo ($msg)) {
            New-Item -ItemType Directory -Force -Path "$path"
        }
        else {
            return $false
        }
    }
    return $true
}

#
# Returns a list of file objects of potential BI Publisher compressed archived in a Path
#
function GetAllBIPExports {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [Alias("Path")]
        [string]$pPath
    )
  
    $gci = Get-ChildItem -Path "$pPath" -Include ($FilterItems) -Recurse
    return $gci
}

#
# Remove BI Publisher Internal files if requested.
# Those files are critical for BI Publisher in the catalog but not necessary for versioning or comparing purpose.
#
function RemoveBIPUselessFiles {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [Alias("Path")]
        [string]$pPath
    )
   
    # Filtered Removal
    if ($PurgeItems.Count -gt 0) {
        Remove-Item -Path "$pPath/*" -Include $PurgeItems -Force
    }
}

#
# Purge Internal BIP Files inside Folders with recursion
#
function RemoveBIPUselessInternalFiles {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [Alias("Path")]
        [string]$pPath
    )

    # Filtered Removal
    if ($PurgeInternals.Count -gt 0) {
        Get-ChildItem -Path "$pPath" -Include $PurgeInternals -Recurse | Remove-Item -Force
    }
}


<#  
    .SYNOPSIS
    Parses a DataModel Metadata and export SQL Statements
    .DESCRIPTION
    This function takes loads a DataModel XML Metadata file and parses the file looking
    for all thepossible SQL queries and saves each found occurence inside a seperate .sql file.
#>
function ParseForSQLStmts ($file, $path) {
    [System.Xml.XmlDocument] $xdoc = new-object System.Xml.XmlDocument
    $xfile = resolve-path($file)
    $xdoc.load($xfile)

    GetSQLStatements $xdoc $path "/ns:dataModel/ns:eventTriggers/ns:eventTrigger[ns:language='SQL']/ns:query" "Trigger"
    GetSQLStatements $xdoc $path "/ns:dataModel/ns:dataSets/ns:dataSet/ns:sql" "DataSet"
    GetSQLStatements $xdoc $path "/ns:dataModel/ns:bursting/ns:burst/ns:dataSet/ns:sql" "Bursting"
    GetSQLStatements $xdoc $path "/ns:dataModel/ns:valueSets/ns:valueSet/ns:sql" "ValueSet"
}

<#  
    .SYNOPSIS
    Generic SQL Statement Extractor
    .DESCRIPTION
    This function prepares a XML Namespace Manager to properly parse the DataModel Metadata and
    parses the given XML file for the XML Selector pattern.
    Each found occurence's content is extracted and saved to a new file.
#>
function GetSQLStatements ($xdoc, $path, $xPathSelector, $sqlType) {

    # Need Namespaces as the XML has some
    $nsmgr = New-Object System.Xml.XmlNamespaceManager $xdoc.NameTable
    $nsmgr.AddNamespace("ns", $xdoc.DocumentElement.NamespaceURI)
    $nsmgr.AddNamespace('xdm', "http://xmlns.oracle.com/oxp/xmlp")
    $nsmgr.AddNamespace('xsd', "http://wwww.w3.org/2001/XMLSchema")

    Write-Host " - Parsing SQL $sqlType ..."
    $xnodes = $xdoc.selectnodes($xPathSelector, $nsmgr)
    # burst name="BURSTING_MAIL"
    foreach ($xnode in $xnodes) {
        # $name = $xnode.
        $name = (GetSQLFilenameFromNode $xnode $sqlType)
        Write-Host "    + BRS Name: $($name)"

        #
        $sqlFilename = Join-Path $path $name
        SaveDOMNodeToFile $xnode.InnerText $sqlFilename    
    }
}

# 
# Get SQL Filename depending on SQL Type.
# This is to avoid file collision as inside a same type, names are distinct but never accross SQL from various nature.
#
function GetSQLFilenameFromNode ($xnode, $sqlType) {
    $name = ""
    # Different parsing depending on the Node.
    # Completly dependant onn BI Publisher Implementation of the DataModel Structure
    switch ($sqlType) {
        DataSet {
            $name = ($xnode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "name") } | ForEach-Object { $_.Value })
            $name = "DS_" + $name + ".sql"
        }
        ValueSet {
            $name = ($xnode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "id") } | ForEach-Object { $_.Value })
            $name = "LOV_" + $name + ".sql"
        }
        Bursting {
            $name = ($xnode.ParentNode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "name") } | ForEach-Object { $_.Value })
            $name = "BRS_" + $name + ".sql"
        }
        Trigger {
            $name = ($xnode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "name") } | ForEach-Object { $_.Value })
            $name = "TRG_" + $name + ".sql"
        }               
    }
    return $name
}

#
# Save XML Node to File
#
function SaveDOMNodeToFile($xdom, $file) {
    # Get
    $Content = $xdom
    [System.IO.File]::WriteAllLines($file, $Content)
}

#
# Handle Objects Exports
# ".xdmz",".xdoz",".xssz",".xsbz"
#
function UnarchiveBIPItemExports ([System.IO.FileInfo[]]$exportFiles) {
    $exportFiles | Where-Object { ($FilterItems -contains "*$($_.Extension)") } | ForEach-Object {
        $ext = "*$($_.Extension)"
        if ($ext -eq "*.xdrz") {
            UnarchiveBIPFolderExports $_
        }
        else {

            # Kept because rework coming
            $extractDir = Resolve-Path $_.FullName
            Write-Host "Extracting Item $($extractDir)"

            # Rename file and add ".zip" to avoid collision and keep track of old/new names    
            $oldFullName = $_.FullName
            $newFullName = $_.FullName + ".tmp"
            $oldName = $_.Name
            $newName = $_.Name + ".tmp"
            Write-Host " - Rename to $($newName)"
            Rename-Item -Path "$oldFullName" -NewName "$newFullName"

            # 7zip Decompression
            #    aoa - Mode Overwrite Mode (Overwrite All existing files without prompt)
            #    bb0 - Set output log level
            #    bd  - Disable progress indicator
            #    o   - Set output directory
            Write-Host " - Extract to $($oldName)"
            # & "$ZIP_CMDPATH" "x" "$newFullName" "-aoa" "-o$oldFullName" "-bd" "-bb2" "-bsp1" "-bse2" "-bso1"
            & "$ZIP_CMDPATH" "x" "$newFullName" "-aoa" "-o$oldFullName" "-bd" "-bb0" "-bsp0" "-bse2" "-bso0" $Exclude7ZipArgs

            # Remove extracted archive
            Write-Host " - Remove $($newName)"
            Remove-Item "$newFullName"

            # Proceed to Item Folder Purge
            # Remove all BIP useless Elements
            # NOTE : Should no be useless with 7zip initial filter
            # if ($DoPurgeItems) {
            #   Write-Host " - Purge Item Folder"
            #   RemoveBIPUselessFiles -Path "$oldFullName"
            # }

            # Post Process Item
            if ($_.Extension -eq ".xdmz") {
                if ($ExtractSQL) {
                    $dmPath = Join-Path $oldFullName "_datamodel.xdm"
                    ParseForSQLStmts $dmPath $oldFullName
                } 
            }
        }
    }
}

#
# Unarchives a BI Publisher folder archive and recursively all it's contained BIP archived items.
# Finishes by a folder cleanup of requested Internal files.
#
function UnarchiveBIPFolderExports ([System.IO.FileInfo[]]$exportFiles) {
    # Unarchive Folders
    $exportFiles | Where-Object { ($_.Extension -eq ".xdrz") } | ForEach-Object {
        # Kept because rework coming
        $extractDir = Join-Path $global:Settings.Arguments.WorkDirPath $_.Name
        Write-Host "Extracting Folder $($extractDir)"

        # 7zip Decompression
        #    aoa - Mode Overwrite Mode (Overwrite All existing files without prompt)
        #    bb0 - Set output log level
        #    bd  - Disable progress indicator
        #    o   - Set output directory    
        # & "$ZIP_CMDPATH" "x" $_.fullname "-aoa" "-o$extractDir" "-bd" "-bb2" "-bsp1" "-bse2" "-bso1"
        & "$ZIP_CMDPATH" "x" $_.Fullname "-aoa" "-o$extractDir" "-bd" "-bb0" "-bsp0" "-bse2" "-bso0" $Exclude7ZipArgs

        # Unarchive potential Items that were contained in the Folder
        $gci = GetAllBIPExports -Path $extractDir
        UnarchiveBIPItemExports $gci

        # NOTE : Should no be useless with 7zip initial filter
        # Cleanup
        # if ($DoPurgeInternals) {
        #   Write-Host "Purge Folders"
        #   RemoveBIPUselessInternalFiles $extractDir
        # }
    }
}

#
#
# From 7Zip Documentation
#        -x[<recurse_type>]<file_ref>
#        <recurse_type> ::= r[- | 0]
#        <file_ref> ::= @{listfile} | !{wildcard}
#
Function FormatExclude7ZipToArgs ($exclude7zip) {
    $exclude7zipArgs = @()
    foreach ($pattern in $exclude7zip) {
        $exclude7zipArgs += ("-xr!" + $pattern)
    }
    return $exclude7zipArgs
}

#
# Main
#
try {

    $global:Settings = @{}

    # TODO : Wrong way to handle. To rework
    try {
        $global:Settings = loadConfiguration "conf.xml"
    }
    catch {
        Write-Host "No configuration to load. Using arguments."

        # Check for parameter validity
        $_rootMsg = "folder does not exist, do you wish to create it? [Y/N]"
        $_wrkMsg = "Work " + $_rootMsg
        $_inMsg = "Input " + $_rootMsg
        $_outMsg = "Output " + $_rootMsg

        if (!(CheckFolderAndCreate "$pInputDirPath" "$_inMsg")) {
            Write-Error -Message "Folder not created. Process stopped !" -ErrorAction Stop
        }
        if (!(CheckFolderAndCreate "$pOutputDirPath" "$_outMsg")) {
            Write-Error -Message "Folder not created. Process stopped !" -ErrorAction Stop
        }
        if (!(CheckFolderAndCreate "$pWorkDirPath" "$_wrkMsg")) {  
            Write-Error -Message "Folder not created. Process stopped !" -ErrorAction Stop
        } 

        # Constants / Globals
        $GUID = [guid]::NewGuid()
        # Folder Information
        $global:Settings.Arguments.InputDirPath = $pInputDirPath
        $global:Settings.Arguments.OutputDirPath = $pOutputDirPath
        $global:Settings.Arguments.WorkDirPath = $pWorkDirPath
        # Extraction flags
        $global:Settings.Arguments.OmitThumbnails = $pOmitThumbnails.IsPresent
        $global:Settings.Arguments.OmitDataSamples = $pOmitDataSamples.IsPresent
        $global:Settings.Arguments.OmitSecurity = $pOmitSecurity.IsPresent 
        $global:Settings.Arguments.OmitMetadata = $pOmitMetadata.IsPresent
        $global:Settings.Arguments.OmitReports = $pOmitReports.IsPresent
        $global:Settings.Arguments.OmitDataModels = $pOmitDataModels.IsPresent
        $global:Settings.Arguments.OmitSubtemplates = $pOmitSubtemplates.IsPresent
        $global:Settings.Arguments.OmitTranslations = $pOmitTranslations.IsPresent
        $global:Settings.Arguments.OmitStyleTemplates = $pOmitStyleTemplates.IsPresent
        # Advanced options
        $global:Settings.Arguments.ExtractSQL = $pExtractSQL.IsPresent

    }
    
    #
    # Data Model      .xdmz      
    # Report          .xdoz  
    # Style Template  .xssz  
    # Subtemplate     .xsbz
    # Pre-Compute Item Archives Filters
    $ExcludeItems = @()
    $FilterItems += @("*.xdrz", "*.xdoz", "*.xdmz", "*.xdmz", "*.xsbz", "*.xliff", "*.xssz")

    if ($global:Settings.Arguments.OmitReports) {
        $ExcludeItems += "*.xdoz"
    }
    if ($global:Settings.Arguments.OmitDataModels) {
        $ExcludeItems += "*.xdmz"
    }
    if ($global:Settings.Arguments.OmitSubtemplates) {
        $ExcludeItems += "*.xsbz"
    }  
    if ($global:Settings.Arguments.OmitTranslations) {
        $ExcludeItems += "*.xliff"
    } 
    if ($global:Settings.Arguments.OmitStyleTemplates) {
        $ExcludeItems += "*.xssz"
    }   


    # Pre-Compute Item advanced Filters
    if ($global:Settings.Arguments.OmitThumbnails) {
        $ExcludeItems += "*.png"
    }
    if ($global:Settings.Arguments.OmitDataSamples) {
        $ExcludeItems += "sample.xml"
    } 
    if ($global:Settings.Arguments.OmitSecurity) {
        $ExcludeItems += "~security.Sec"
    }  
    if ($global:Settings.Arguments.OmitMetadata) {
        $ExcludeItems += "~metadata.meta"
    }  

    # Pre-Compute 7Zip exclude patterns
    $Exclude7ZipArgs = @()
    $Exclude7ZipArgs = FormatExclude7ZipToArgs $ExcludeItems

    # DEBUG
    Write-Host "--------------------------------------------------------------------------------"
    Write-Host "BI Publisher Catalog Extractor"
    Write-Host "--------------------------------------------------------------------------------"
    Write-Host "Input Directory     : $($global:Settings.Arguments.InputDirPath)"
    Write-Host "Output Directory    : $($global:Settings.Arguments.OutputDirPath)"
    Write-Host "Work Directory      : $($global:Settings.Arguments.WorkDirPath)"
    Write-Host "Extract SQL         : $($global:Settings.Arguments.ExtractSQL)"
    Write-Host "==== Extraction ===="    
    Write-Host "Omit Reports        : $($global:Settings.Arguments.OmitReports)"
    Write-Host "Omit DataModels     : $($global:Settings.Arguments.OmitDataModels)"
    Write-Host "Omit Subtemplates   : $($global:Settings.Arguments.OmitSubtemplates)"
    Write-Host "Omit Translations   : $($global:Settings.Arguments.OmitTranslations)"
    Write-Host "Omit Styles         : $($global:Settings.Arguments.OmitStyleTemplates)"  
    Write-Host "==== Advanced ======"
    Write-Host "Omit Security       : $($global:Settings.Arguments.OmitSecurity)"
    Write-Host "Omit Thumbnails     : $($global:Settings.Arguments.OmitThumbnails)"
    Write-Host "Omit RPT Metadata   : $($global:Settings.Arguments.OmitMetadata)"
    Write-Host "Omit DM DataSamples : $($global:Settings.Arguments.OmitDataSamples)"  
    Write-Host "==== Debug ========="
    Write-Host "GUID                : $GUID"
    Write-Host "--------------------------------------------------------------------------------"

    # ---------------------------------------------------------------------------
    # BI Publisher Archives Extensions
    # Folder          .xdrz
    # Data Model      .xdmz      
    # Report          .xdoz  
    # Style Template  .xssz  
    # Subtemplate     .xsbz
    # ---------------------------------------------------------------------------

    # Cleanup Temporary Folder
    Get-ChildItem -LiteralPath "$($global:Settings.Arguments.WorkDirPath)" -Directory -Recurse | ForEach-Object {
        Remove-Item -Path $_.FullName -Force -Recurse
    }

    # Get the candidate archive files and Unarchive / Cleanup
    $ExportFiles = Get-ChildItem -Path "$($global:Settings.Arguments.InputDirPath)" -Include ($FilterItems) -Recurse
    UnarchiveBIPItemExports $ExportFiles

    Write-Host "--------------------------------------------------------------------------------"
    # Exit Routine
    exit 0
}
catch {
    Write-Error $_
}

exit 0