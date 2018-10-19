###############################################################################
#
#                        BI Publisher Archive Manager
#
#
# Author  : Francois CABANNES
# Version : 1.2.0
# Date    : 19/10/2018
#
##############################################################################
# 
# 1.2.0 : Major rework of Config File and Features - No backward compatibility
# 
###############################################################################
#
# TODO : Code : Real error handling
# TODO : Code : Cleaup exit status
# TODO : Feature : Connect to Remote BIP ad download archives
# TODO : Feature : Connect with Certs
# TODO : Feature : Parametrage Unarchiver and Archive pattern
#
###############################################################################

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

# Set-PSDebug -Trace 1

# BIP Archive Types
Set-Variable BIP_XDR -value ".xdrz"             # Folders
Set-Variable BIP_XDM -value ".xdmz"             # Data Model
Set-Variable BIP_XDO -value ".xdoz"             # Report
Set-Variable BIP_XSS -value ".xssz"             # Style Template
Set-Variable BIP_XSB -value ".xsbz"             # Subtemplate
Set-Variable BIP_XLF -value ".xliff"            # XLIFF Translation

# Global Objects
Set-Variable BIP_SEC -value "~security.sec"     # Security filename
Set-Variable BIP_MET -value "~metadata.meta"    # Metadata filename

# DataModel Objects
Set-Variable DM_XDM -value "_datamodel.xdm"    # Thumbnail filename
Set-Variable DM_XLS -value ".xlsdata"          # XLS DataSet Source
Set-Variable DM_XLSX -value ".xlsxdata"        # XLSX DataSet Source
Set-Variable DM_CSV -value ".csvdata"          # CSV DataSet Source
Set-Variable DM_XML -value "_xdo_local.*.xml"  # XML DataSet Source
Set-Variable DM_XSD -value ".xsd"              # XSD DataSet Source
Set-Variable DM_SMP -value "sample.xml"        # Sample filename

# Report Objects
Set-Variable RPT_XDO -value "_report.xdo"       # Report Metadata
Set-Variable RPT_CFG -value "xdo.cfg"           # Report Property File
Set-Variable RPT_THB -value ".png"              # Thumbnail filename
Set-Variable TPL_XLS -value ".xls"              # Excel 97 Template
Set-Variable TPL_XSL -value ".xsl"              # XSL + XSLFO Templates
Set-Variable TPL_PDF -value ".pdf"              # PDF Template
Set-Variable TPL_RTF -value ".rtf"              # RTF Template
Set-Variable TPL_XPT -value ".xpt"              # Interactive Report Template
Set-Variable TPL_RTF -value ".xlf"              # XLIFF Translation

# Data Model SQL Prefixes
Set-Variable BIP_SQL -value ".sql"              # SQL Filename extension
Set-Variable SQL_DM_PREFIX -value "DM_"         # Prefix for DataModel SQL filename 
Set-Variable SQL_LOV_PREFIX -value "LOV_"       # Prefix for ValueSets SQL filename 
Set-Variable SQL_TRG_PREFIX -value "TRG_"       # Prefix for Trigger SQL filename 
Set-Variable SQL_BRS_PREFIX -value "BRS_"       # Prefix for Bursting SQL filename 

Set-Variable DEFAULT_CFG_FILE -value "conf.xml"

#
#
#
function loadConfiguration ($pConfigFile) {
    # Import email settings from config file
    [xml]$ConfigFile = Get-Content ($pConfigFile)
    $Settings = @{
        ZipSettings = @{
            Path   = $ConfigFile.Settings.ZipSettings.Path
            Extras = $ConfigFile.Settings.ZipSettings.Extras
        }
        AppSettings = @{
            Certs = $ConfigFile.Settings.AppSettings.Certs
        }
        Ext         = @{
            InputDirPath       = $ConfigFile.Settings.Extractor.InputDirPath
            OutputDirPath      = $ConfigFile.Settings.Extractor.OutputDirPath
            WorkDirPath        = $ConfigFile.Settings.Extractor.WorkDirPath
            OmitSecurity       = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitSecurity)
            OmitMetadata       = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitMetadata)
            OmitReports        = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitReports)
            OmitDataModels     = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitDataModels)
            OmitSubtemplates   = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitSubtemplates)
            OmitTranslations   = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitTranslations)
            OmitStyleTemplates = [System.Convert]::ToBoolean($ConfigFile.Settings.Extractor.OmitStyleTemplates)
        }
        Dm          = @{
            OmitMetadata  = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitMetadata)
            OmitExcel     = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitExcel)
            OmitExcel2007 = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitExcel2007)
            OmitCSV       = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitCSV)
            OmitXML       = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitXML)
            OmitSchema    = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitSchema)
            OmitSamples   = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.OmitSamples)
            ExtractSQL    = [System.Convert]::ToBoolean($ConfigFile.Settings.DataModel.ExtractSQL)
        }
        Rpt         = @{
            OmitMetadata   = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitMetadata)
            OmitProperties = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitProperties)
            OmitThumbnails = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitThumbnails)
            OmitExcel      = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitExcel)
            OmitXSL        = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitXSL)
            OmitXSLFO      = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitXSLFO)
            OmitPDF        = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitPDF)
            OmitRTF        = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitRTF)
            OmitETEXT      = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitETEXT)
            OmitXPT        = [System.Convert]::ToBoolean($ConfigFile.Settings.Report.OmitXPT)
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
#
#
function RemoveEmptyFolders {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [Alias("Path")]
        [string]$pPath,

        [Parameter(Mandatory = $False)]
        [Alias("SkipRoot")]
        [bool]$pSkipRoot = $True       
    )
    #
    # https://stackoverflow.com/questions/28631419/how-to-recursively-remove-all-empty-folders-in-powershell
    # Thanks to Kirk Munro
    #
    # Now define a script block that will remove empty folders under
    # a root folder, using tail-recursion to ensure that it only
    # walks the folder tree once. -Force is used to be able to process
    # hidden files/folders as well.

    foreach ($childDirectory in Get-ChildItem -Force -LiteralPath $pPath -Directory) {
        RemoveEmptyFolders -Path $childDirectory.FullName -SkipRoot $False
    }
    $currentChildren = Get-ChildItem -Force -LiteralPath $pPath
    $isEmpty = ($currentChildren -eq $null)
    if ($isEmpty -And !$pSkipRoot) {
        Write-Verbose "Removing empty folder at path '${pPath}'." -Verbose
        Remove-Item -Force -LiteralPath $pPath
    }
}

#
# Remove BI Publisher Internal files if requested.
# Those files are critical for BI Publisher in the catalog but not necessary for versioning or comparing purpose.
#
function RemoveBIPUselessFiles {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [Alias("Path")]
        [string]$pPath,

        [Parameter()]
        [Alias("PurgeItems")]
        [string[]]$pPurgeItems
    )
   
    # Filtered Removal
    if ($pPurgeItems.Count -gt 0) {
        Write-Host " - Purging : $($pPurgeItems -join ', ')"
        Remove-Item -Path "$pPath/*" -Include $pPurgeItems -Force
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
        [string]$pPath,

        [Parameter()]
        [Alias("PurgeItems")]
        [string[]]$pPurgeItems
    )
    
    # Filtered Removal
    if ($pPurgeItems.Count -gt 0) {
        Get-ChildItem -Path "$pPath" -File -Include $pPurgeItems -Recurse | Remove-Item -Force -Recurse -Verbose
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
        $filename = (GetSQLFilenameFromNode $xnode $sqlType)
        Write-Host "    + BRS Name: $($filename)"

        SaveDOMNodeToFile $xnode.InnerText $path $filename    
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
            $name = ($SQL_DM_PREFIX + $name + $BIP_SQL)
        }
        ValueSet {
            $name = ($xnode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "id") } | ForEach-Object { $_.Value })
            $name = ($SQL_LOV_PREFIX + $name + $BIP_SQL)
        }
        Bursting {
            $name = ($xnode.ParentNode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "name") } | ForEach-Object { $_.Value })
            $name = ($SQL_BRS_PREFIX + $name + $BIP_SQL)
        }
        Trigger {
            $name = ($xnode.ParentNode.Attributes | Where-Object { ( $_.Name -eq "name") } | ForEach-Object { $_.Value })
            $name = ($SQL_TRG_PREFIX + $name + $BIP_SQL)
        }               
    }
    return $name
}

#
# Save XML Node to a valid Filename.
#
function SaveDOMNodeToFile($xdom, $path, $filename) {   
    $fileFixed = $filename
    [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
        $fileFixed = $fileFixed.replace($_, '.')
    }
    $fileFixed = (Join-Path $path $fileFixed)
    [System.IO.File]::WriteAllLines($fileFixed, $xdom)
}

#
# Handle Objects Exports
#
function UnarchiveBIPItemExports ([System.IO.FileInfo[]]$exportFiles) {
    $exportFiles | Where-Object { ($FilterItems -contains "*$($_.Extension)") } | ForEach-Object {
        $ext = "*$($_.Extension)"
        # Folders
        if ($ext -eq "*$($BIP_XDR)") {
            UnarchiveBIPFolderExports $_
        }
        # All Others
        else {
            $excludeArgs = @()
            $excludeArgs += $ExcludeAlwaysArgs
            $purgeItems = @()           

            # Prepare exclude and purge patterns
            if ($ext -eq "*$($BIP_XDM)") {
                $excludeArgs += $ExcludeDMArgs
                $purgeItems += $PurgeDMItems
            } 
            elseif ($ext -eq "*$($BIP_XDO)") {
                $excludeArgs += $ExcludeRPTArgs
                $purgeItems += $PurgeRPTItems
            }

            $extractDir = Resolve-Path $_.FullName
            Write-Host "Processing $($extractDir)"            

            # Rename file and add ".zip" to avoid collision and keep track of old/new names    
            $oldFullName = $_.FullName
            $newFullName = $_.FullName + ".tmp"
            $oldName = $_.Name
            $newName = $_.Name + ".tmp"
            Write-Host " - Renaming to $($newName)"
            Rename-Item -Path "$oldFullName" -NewName "$newFullName"

            # Warning all 7zip versions do not handle switches for granularity of logging
            # TODO : Inject unzipping pattern in Config File
            # TODO : Will help handle custom unarchivers
            #
            # 7zip Decompression
            #    aoa - Mode Overwrite Mode (Overwrite All existing files without prompt)
            #    bb0 - Set output log level
            #    bd  - Disable progress indicator
            #    o   - Set output directory
            Write-Host " - Extracting to $($oldName)"
            Write-Host " - Skipping : $($excludeArgs -join ', ')"
            # & "$ZIP_CMDPATH" "x" "$newFullName" "-aoa" "-o$oldFullName" "-bd" "-bb2" "-bsp1" "-bse2" "-bso1"
            & "$($global:Settings.ZipSettings.Path)" "x" "$($newFullName)" "-aoa" "-o$($oldFullName)" ($(($global:Settings.ZipSettings.Extras).Split(" "))) ($excludeArgs)

            # Remove temporary archive
            Write-Host " - Removing $($newName)"
            Remove-Item "$newFullName"
            # Post-Process Item depending on Item type then trigger a purge
            if ($_.Extension -eq $BIP_XDM) {
                if ($global:Settings.Dm.ExtractSQL) {
                    $dmPath = Join-Path $oldFullName $DM_XDM
                    ParseForSQLStmts $dmPath $oldFullName
                }                
            }
            RemoveBIPUselessFiles -Path $_.FullName -PurgeItems $purgeItems
        }
    }
}

#
# Unarchives a BI Publisher folder archive and recursively all it's contained BIP archived items.
# Finishes by a folder cleanup of requested Internal files.
#
function UnarchiveBIPFolderExports ([System.IO.FileInfo[]]$exportFiles) {
    # Unarchive Folders
    $exportFiles | Where-Object { ($_.Extension -eq $BIP_XDR) } | ForEach-Object {
        # Kept because rework coming
        $extractDir = Join-Path $global:Settings.Ext.WorkDirPath $_.Name
        Write-Host "Extracting Folder $($extractDir)"

        # 7zip Decompression
        #    aoa - Mode Overwrite Mode (Overwrite All existing files without prompt)
        #    bb0 - Set output log level
        #    bd  - Disable progress indicator
        #    o   - Set output directory    
        # & "$ZIP_CMDPATH" "x" $_.fullname "-aoa" "-o$extractDir" "-bd" "-bb2" "-bsp1" "-bse2" "-bso1"
        & "$($global:Settings.ZipSettings.Path)" "x" "$($_.Fullname)" "-aoa" "-o$($extractDir)" ($(($global:Settings.ZipSettings.Extras).Split(" "))) ($Exclude7ZipArgs)

        # Unarchive potential Items that were contained in the Folder
        $gci = GetAllBIPExports -Path $extractDir
        UnarchiveBIPItemExports $gci
    }
}

#
# 
#
Function FormatExclude7ZipToArgs ($exclude7zip) {
    # From 7Zip Documentation
    #        -x[<recurse_type>]<file_ref>
    #        <recurse_type> ::= r[- | 0]
    #        <file_ref> ::= @{listfile} | !{wildcard}
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

    $global:GUID = [guid]::NewGuid()
    $global:Settings = @{}

    # TODO : Wrong way to handle. To rework ...
    try {
        if ($pConfig -ne "") {
            $confFile = $pConfig
        }
        else {
            $confFile = $DEFAULT_CFG_FILE
        }
        $global:Settings = loadConfiguration $confFile
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

        # TODO : Rewrite completly due to new parameters
        # Folder Information
        $global:Settings.Ext.InputDirPath = $pInputDirPath
        $global:Settings.Ext.OutputDirPath = $pOutputDirPath
        $global:Settings.Ext.WorkDirPath = $pWorkDirPath
        # Extraction flags
        $global:Settings.Ext.OmitThumbnails = $pOmitThumbnails.IsPresent
        $global:Settings.Ext.OmitDataSamples = $pOmitDataSamples.IsPresent
        $global:Settings.Ext.OmitSecurity = $pOmitSecurity.IsPresent 
        $global:Settings.Ext.OmitMetadata = $pOmitMetadata.IsPresent
        $global:Settings.Ext.OmitReports = $pOmitReports.IsPresent
        $global:Settings.Ext.OmitDataModels = $pOmitDataModels.IsPresent
        $global:Settings.Ext.OmitSubtemplates = $pOmitSubtemplates.IsPresent
        $global:Settings.Ext.OmitTranslations = $pOmitTranslations.IsPresent
        $global:Settings.Ext.OmitStyleTemplates = $pOmitStyleTemplates.IsPresent
        # Advanced options
        $global:Settings.Ext.ExtractSQL = $pExtractSQL.IsPresent
    }
        
    # Pre-Compute Item Archives Filters
    $ExcludeItems = @()
    $ExcludeAlwaysItems = @()
    $ExcludeRPTItems = @()
    $ExcludeDMItems = @()
    $PurgeRPTItems = @()
    $PurgeDMItems = @()    

    $FilterItems += @(
        "*$($BIP_XDR)",
        "*$($BIP_XDM)",
        "*$($BIP_XDO)",
        "*$($BIP_XSS)",
        "*$($BIP_XSB)",
        "*$($BIP_XLF)"
    )

    # Archive Excludes
    # Adding * before patterns only
    if ($global:Settings.Ext.OmitReports) {
        $ExcludeItems += "*$($BIP_XDO)"
    }
    if ($global:Settings.Ext.OmitDataModels) {
        $ExcludeItems += "*$($BIP_XDM)"
    }
    if ($global:Settings.Ext.OmitSubtemplates) {
        $ExcludeItems += "*$($BIP_XSB)"
    }  
    if ($global:Settings.Ext.OmitTranslations) {
        $ExcludeItems += "*$($BIP_XLF)"
    } 
    if ($global:Settings.Ext.OmitStyleTemplates) {
        $ExcludeItems += "*$($BIP_XSS)"
    }   

    # Global Excludes
    if ($global:Settings.Ext.OmitSecurity) {
        $ExcludeItems += "$($BIP_SEC)"
        $ExcludeAlwaysItems += "$($BIP_SEC)"
        $ExcludeRPTItems += "$($BIP_SEC)"
        $ExcludeDMItems += "$($BIP_SEC)"            
    }  
    if ($global:Settings.Ext.OmitMetadata) {
        $ExcludeItems += "$($BIP_MET)"
        $ExcludeAlwaysItems += "$($BIP_MET)"
        $ExcludeRPTItems += "$($BIP_MET)"
        $ExcludeDMItems += "$($BIP_MET)"       
    }  

    # Report Excludes
    if ($global:Settings.Rpt.OmitMetadata) {
        $PurgeRPTItems += "$($RPT_XDO)"
    }
    if ($global:Settings.Rpt.OmitProperties) {
        $ExcludeRPTItems += "*$($RPT_CFG)"
    }    
    if ($global:Settings.Rpt.OmitThumbnails) {
        $ExcludeRPTItems += "*$($RPT_THB)"
    }
    if ($global:Settings.Rpt.OmitExcel) {
        $ExcludeRPTItems += "*$($TPL_XLS)"
    }
    if ($global:Settings.Rpt.OmitXSL) {
        $ExcludeRPTItems += "*$($TPL_XSL)"
    }
    # TO IMPLEMENT  - Same time as .XSL ATM
    # if ($global:Settings.Rpt.OmitXSLFO) {
    #     $ExcludeRPTItems += "*$($TPL_XSLFO)"
    # }
    if ($global:Settings.Rpt.OmitPDF) {
        $ExcludeRPTItems += "*$($TPL_PDF)"
    }
    if ($global:Settings.Rpt.OmitRTF) {
        $ExcludeRPTItems += "*$($TPL_RTF)"
    }
    # TO IMPLEMENT - Same time as .RTF ATM
    # if ($global:Settings.Rpt.OmitETEXT) {
    #     $ExcludeRPTItems += "*$($TPL_THB)"
    # }
    if ($global:Settings.Rpt.OmitXPT) {
        $ExcludeRPTItems += "*$($TPL_XPT)"
    }         

    # DataModel Excludes    
    if ($global:Settings.Dm.OmitMetadata) {
        $PurgeDMItems += "$($DM_XDM)"
    } 
    if ($global:Settings.Dm.OmitSamples) {
        $ExcludeDMItems += "$($DM_SMP)"
    }
    if ($global:Settings.Dm.OmitExcel) {
        $ExcludeDMItems += "*$($DM_XLS)"
    } 
    if ($global:Settings.Dm.OmitExcel2007) {
        $ExcludeDMItems += "*$($DM_XLSX)"
    } 
    if ($global:Settings.Dm.OmitCSV) {
        $ExcludeDMItems += "*$($DM_CSV)"
    } 
    if ($global:Settings.Dm.OmitXML) {
        $ExcludeDMItems += "*$($DM_XML)"
    } 
    if ($global:Settings.Dm.OmitSchema) {
        $ExcludeDMItems += "*$($DM_XSD)"
    } 

    # Pre-Compute 7Zip exclude patterns
    $ExcludeAlwaysArgs = FormatExclude7ZipToArgs $ExcludeAlwaysItems
    $ExcludeRPTArgs = FormatExclude7ZipToArgs $ExcludeRPTItems
    $ExcludeDMArgs = FormatExclude7ZipToArgs $ExcludeDMItems    
    $Exclude7ZipArgs = FormatExclude7ZipToArgs $ExcludeItems

    # DEBUG
    Write-Host "--------------------------------------------------------------------------------"
    Write-Host "BI Publisher Catalog Extractor"
    Write-Host "--------------------------------------------------------------------------------"
    Write-Host "Input Directory     : $($global:Settings.Ext.InputDirPath)"
    Write-Host "Output Directory    : $($global:Settings.Ext.OutputDirPath)"
    Write-Host "Work Directory      : $($global:Settings.Ext.WorkDirPath)"
    
    Write-Host "==== Extraction ===="        
    Write-Host "Omit Reports        : $($global:Settings.Ext.OmitReports)"
    Write-Host "Omit DataModels     : $($global:Settings.Ext.OmitDataModels)"
    Write-Host "Omit Subtemplates   : $($global:Settings.Ext.OmitSubtemplates)"
    Write-Host "Omit Translations   : $($global:Settings.Ext.OmitTranslations)"
    Write-Host "Omit Styles         : $($global:Settings.Ext.OmitStyleTemplates)"  

    Write-Host "==== Advanced ======"
    Write-Host "Omit Security       : $($global:Settings.Ext.OmitSecurity)"
    Write-Host "Omit Metadata       : $($global:Settings.Ext.OmitMetadata)"    

    Write-Host "==== DataModels ===="
    Write-Host "Omit Metadata       : $($global:Settings.Dm.OmitMetadata)"  
    Write-Host "Omit CSV            : $($global:Settings.Dm.OmitCSV)"
    Write-Host "Omit XML            : $($global:Settings.Dm.OmitXML)"
    Write-Host "Omit Excel          : $($global:Settings.Dm.OmitExcel)"
    Write-Host "Omit Excel 2007     : $($global:Settings.Dm.OmitExcel2007)"
    Write-Host "Omit Schema         : $($global:Settings.Dm.OmitSchema)"
    Write-Host "Omit Samples        : $($global:Settings.Dm.OmitSamples)"
    Write-Host "ExtractSQL          : $($global:Settings.Dm.ExtractSQL)"

    Write-Host "==== Reports ======="        
    Write-Host "Omit Metadata       : $($global:Settings.Rpt.OmitMetadata)"
    Write-Host "Omit Properties     : $($global:Settings.Rpt.OmitProperties)"
    Write-Host "Omit Thumbnails     : $($global:Settings.Rpt.OmitThumbnails)"
    Write-Host "Omit Excel 97       : $($global:Settings.Rpt.OmitExcel)"
    Write-Host "Omit XSL            : $($global:Settings.Rpt.OmitXSL)"
    Write-Host "Omit XSL-FO         : $($global:Settings.Rpt.OmitXSLFO)"
    Write-Host "Omit PDF            : $($global:Settings.Rpt.OmitPDF)"
    Write-Host "Omit RTF            : $($global:Settings.Rpt.OmitRTF)"
    Write-Host "Omit ETEXT          : $($global:Settings.Rpt.OmitETEXT)"
    Write-Host "Omit Inter. Report  : $($global:Settings.Rpt.OmitXPT)"    

    Write-Host "==== Debug ========="
    Write-Host "GUID                : $($global:GUID)"
    Write-Host "--------------------------------------------------------------------------------"

    # Cleanup Temporary Folder
    Get-ChildItem -LiteralPath "$($global:Settings.Ext.WorkDirPath)" -Directory -Recurse | ForEach-Object {
        Remove-Item -Path $_.FullName -Force -Recurse
    }

    # Get the candidate archive files and process
    $ExportFiles = Get-ChildItem -Path "$($global:Settings.Ext.InputDirPath)" -File -Include ($FilterItems) -Recurse
    Write-Host ">> Processing archives"
    UnarchiveBIPItemExports $ExportFiles
    Write-Host ">> Removing internal files ..."
    RemoveBIPUselessInternalFiles  "$($global:Settings.Ext.WorkDirPath)" $ExcludeAlwaysItems
    Write-Host ">> Removing empty folders ..."
    RemoveEmptyFolders "$($global:Settings.Ext.WorkDirPath)\"

    Write-Host "--------------------------------------------------------------------------------"

    exit 0
}
catch {
    # TODO : Need real error handling all over the place
    Write-Error $_
}

exit 0