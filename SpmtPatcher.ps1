<#
    Notes:
    - have SPMT intalled
    - put 0Harmony.dll in the script directory
    - run the script
    - restart PowerShell process after modifications
    - no red error messages = success
#>

function LoadSpmtPowerShellModule() {
    if (Get-Module Microsoft.SharePoint.MigrationTool.PowerShell) {
        Write-Host "SPMT PowerShell module is loaded, good"
        return
    }

    $currentLocation = Get-Location
    $spmtModulePath = "$($env:UserProfile)\Documents\WindowsPowerShell\Modules\Microsoft.SharePoint.MigrationTool.PowerShell"
    Write-Host "Importing SPMT PowerShell module from '$spmtModulePath'"
    if (-not (Get-Item $spmtModulePath -ErrorAction SilentlyContinue).Exists) {
        throw "SharePoint Migration Toolkit PowerShell Module expected but not found in '$spmtModulePath'"

    }
    $null = Set-Location $spmtModulePath
    $null = Import-Module "Microsoft.SharePoint.MigrationTool.PowerShell"
    $null = Set-Location $currentLocation

    if (Get-Module Microsoft.SharePoint.MigrationTool.PowerShell) {
        Write-Host "SPMT PowerShell module loaded" -ForegroundColor Green
    }
    else {
        throw "SPMT PowerShell module not loaded, cannot continue"
    }
}

function LoadHarmony() {
    $harmonyType = ([System.Management.Automation.PSTypeName]'HarmonyLib.Harmony').Type
    if ($harmonyType) {
        Write-Host "Harmony is already loaded, good"
        return
    }

    $harmonyDllPath = "0Harmony.dll"
    if (-not (Get-Item $harmonyDllPath -ErrorAction SilentlyContinue).Exists) {
        if ($PSScriptRoot) {
            $harmonyDllPath = "$PSScriptRoot\0Harmony.dll"
            if (-not (Get-Item $harmonyDllPath -ErrorAction SilentlyContinue).Exists) {
                throw "Did not find Harmony DLL at path '$harmonyDllPath'"
            }
        }
        else {
            Write-Host "Cannot search in script root since you are only running a selection of the script file" -ForegroundColor Yellow
            throw "Did not find Harmony DLL at path '$harmonyDllPath'"
        }
    }
    Write-Host "Found Harmony DLL at path '$harmonyDllPath'" -ForegroundColor Green

    $null = Add-Type -LiteralPath $harmonyDllPath
    $harmonyType = ([System.Management.Automation.PSTypeName]'HarmonyLib.Harmony').Type
    if ($harmonyType) {
        $harmonyTargetFramework = [HarmonyLib.Harmony].Assembly.CustomAttributes |
        Where-Object { $_.AttributeType.Name -eq "TargetFrameworkAttribute" } |
        Select-Object -ExpandProperty ConstructorArguments |
        Select-Object -ExpandProperty value

        Write-Host "Harmony is loaded; version: $([HarmonyLib.Harmony].Assembly.GetName().Version.ToString()); ImageRuntimeVersion: $([HarmonyLib.Harmony].Assembly.ImageRuntimeVersion); TargetFramework: $($harmonyTargetFramework)" -ForegroundColor Green
    }
    else {
        throw "Harmony not loaded"
    }
}

function CreateCustomTypeForHarmony() {
    $Version = 2

    $spmtModificationsType = ([System.Management.Automation.PSTypeName]'SpmtModifications').Type
    # check if version matches or if we got an old type in memory
    if ($spmtModificationsType) {
        if ([SpmtModifications]::Version -ne $Version) {
            throw "Detected old SpmtModifications type in memory. This can happen after code modifications. Please restart the PowerShell process and run the script again."
        }
        Write-Host "Found existing custom SPMT modification type, good"
        return
    }

    $source = Get-Content -Path "$PSScriptRoot\SpmtPatch.cs" -Encoding UTF8 -Raw
    $spmtModule = Get-Module "Microsoft.SharePoint.MigrationTool.PowerShell"
    $spmtModulePath = (Get-Item $spmtModule.Path).Directory.FullName

    Write-Host "Loading types from SPMT module path '$spmtModulePath'"
    $null = Add-Type -Path "$spmtModulePath\microsoft.sharepoint.migrationtool.migrationlib.dll"
    $null = Add-Type -Path "$spmtModulePath\Microsoft.Identity.Client.dll"
    $null = Add-Type `
        -TypeDefinition $source `
        -ReferencedAssemblies "mscorlib",
            "System",
            "System.Core",
            "System.Management.Automation",
            "System.Data",
            "System.Data.DataSetExtensions",
            "System.Drawing",
            "System.IdentityModel",
            "System.Net.Http",
            "System.Windows.Forms",
            "System.Xml",
            "System.Xml.Linq",
            "System.Numerics",
            "System.Runtime.Serialization",
            "Microsoft.Extensions.Configuration",
            "Microsoft.Extensions.Configuration.Abstractions",
            "Microsoft.Extensions.DependencyInjection",
            "Microsoft.Extensions.DependencyInjection.Abstractions",
            "Microsoft.Extensions.Logging",
            "Microsoft.Extensions.Logging.Abstractions",
            "Microsoft.Extensions.Options",
            "Microsoft.Extensions.Primitives",
            "Microsoft.CSharp",
            "$spmtModulePath\Microsoft.Identity.Client.dll",
            "$spmtModulePath\microsoft.sharepoint.client.dll",
            "$spmtModulePath\microsoft.sharepoint.client.runtime.dll",
            "$spmtModulePath\microsoft.sharepoint.migrationtool.migrationlib.dll",
            "$spmtModulePath\microsoft.sharepoint.migration.common.dll",
            "$spmtModulePath\microsoft.sharepoint.migrationtool.powershell.dll",
            "$spmtModulePath\Microsoft.Identity.Client.Extensions.Msal.dll" `
        -IgnoreWarnings `
        -ErrorAction Continue

    $spmtModificationsType = ([System.Management.Automation.PSTypeName]'SpmtModifications').Type
    if (-not $spmtModificationsType) {
        throw "Could not create SpmtModifications type"
    }

    Write-Host "Created custom SPMT modification type" -ForegroundColor Green
}

function PatchSpmt() {
    if ($Global:patched) {
        Write-Host "SPMT methods are already patched, good"
        return
    }

    $patches = @(
        @{
            "origName" = "ValidateListMetaInfo"
            "methodOrig" = [Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract].GetMethod("ValidateListMetaInfo", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::NonPublic)
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("ValidateListMetaInfoPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "NeedSkipInternalView"
            "methodOrig" = [Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.ViewMigrator].GetMethod("NeedSkipInternalView", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::NonPublic)
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("NeedSkipInternalViewPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "IsFilteredList"
            "methodOrig" = [Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.SPODocumentAcquirer].GetMethod("IsFilteredList", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::NonPublic)
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("IsFilteredListPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "CheckListTemplateMatch"
            "methodOrig" = [Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaUtilities].GetMethod("CheckListTemplateMatch", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPrefix" = [SpmtModifications].GetMethod("CheckListTemplateMatchPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "UpsertView"
            "methodOrig" = [Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.ViewMigrator].GetMethod("UpsertView", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::Public)
            "methodPrefix" = [SpmtModifications].GetMethod("UpsertViewPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "CheckOnPremSiteAccessibility"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.PowerShell.MigrationFeasibilityChecker"), "CheckOnPremSiteAccessibility")
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("CheckOnPremSiteAccessibilityPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "GetSPOnPremDocumentList"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.SPODocumentAcquirer"), "GetSPOnPremDocumentList")
            "methodPrefix" = [SpmtModifications].GetMethod("GetSPOnPremDocumentListPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "SPODocumentAcquirer.CreateContext"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.SPODocumentAcquirer"), "CreateContext")
            "methodPrefix" = [SpmtModifications].GetMethod("SPODocumentAcquirer_CreateContextPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "SharePointContext.CreateContext"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.SharePointContext"), "CreateContext")
            "methodPrefix" = [SpmtModifications].GetMethod("SharePointContext_CreateContextPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "CreateClientContext"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.Common.MIGUtilities"), "CreateClientContext")
            "methodPrefix" = [SpmtModifications].GetMethod("CreateClientContextPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = [SpmtModifications].GetMethod("CreateClientContextPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "InitAuthorizationHeader"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.MigSPOnlineClientContext"), "InitAuthorizationHeader")
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("InitAuthorizationHeaderPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "AcquireAccessToken"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.Common.AADContext"), "AcquireAccessToken", @([string], [string], [System.Security.SecureString], [bool]))
            "methodPrefix" = [SpmtModifications].GetMethod("AcquireAccessTokenPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = [SpmtModifications].GetMethod("AcquireAccessTokenPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "IsSourceVersionSupported"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.Common.MIGUtilities"), "IsSourceVersionSupported", [int])
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("IsSourceVersionSupportedPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "IsSourceVersionSupported2"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.Common.MIGUtilities"), "IsSourceVersionSupported", [Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.IMigSPEnvironment])
            "methodPrefix" = $null
            "methodPostfix" = [SpmtModifications].GetMethod("IsSourceVersionSupported2Postfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "AssessSpSiteAdmin"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SourceProviderAbstract"), "AssessSpSiteAdmin")
            "methodPrefix" = [SpmtModifications].GetMethod("AssessSpSiteAdminPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = [SpmtModifications].GetMethod("AssessSpSiteAdminPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        }
        @{
            "origName" = "SharePointContext.ctor"
            "methodOrig" = [HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.SharePointContext").GetMember(".ctor")[0]
            "methodPrefix" = [SpmtModifications].GetMethod("SharePointContext_ctorPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "OpenBinaryDirect"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.FileProxy"), "OpenBinaryDirect")
            "methodPrefix" = [SpmtModifications].GetMethod("OpenBinaryDirectPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "GetHttpResponse"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.MigrationLib.Common.Net.RestApiSDK"), "GetHttpResponse")
            "methodPrefix" = [SpmtModifications].GetMethod("GetHttpResponsePrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        @{
            "origName" = "BeginProcessing"
            "methodOrig" = [HarmonyLib.AccessTools]::Method([HarmonyLib.AccessTools]::TypeByName("Microsoft.SharePoint.MigrationTool.PowerShell.AddSPMTTask"), "BeginProcessing")
            "methodPrefix" = [SpmtModifications].GetMethod("ProcessRecordPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
    )
    # patch all logging methods as well
    $logTypes = "Information", "Error", "Warning", "Exception", "Debug", "Verbose"
    foreach ($logType in $logTypes) {
        $patch = @{
            "origName" = "Log$($logType)"
            "methodOrig" = [Microsoft.SharePoint.MigrationTool.MigrationLib.Common.MigObjectAbstract].GetMethod("Log$($logType)", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::Public)
            "methodPrefix" = [SpmtModifications].GetMethod("Log$($logType)Prefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
            "methodPostfix" = $null
        }
        $patches += $patch
    }

    # a bit more work for getter
    $prop = [Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract].GetProperty("SupportedListTemplates", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    $getter = $prop.GetAccessors()
    $patch = @{
        "origName" = "GetSupportedListTemplates"
        "methodOrig" = $getter[0]
        "methodPrefix" = $null
        "methodPostfix" = [SpmtModifications].GetMethod("GetSupportedListTemplatesPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    }
    $patches += $patch

    # apply patches
    foreach ($patch in $patches)
    {
        $harmony = New-Object HarmonyLib.Harmony -ArgumentList "com.heinrich.spmtredirects"

        $prefix = $null
        $postfix = $null
        if ($patch.methodPrefix) {
            $prefix = (New-Object HarmonyLib.HarmonyMethod -ArgumentList $patch.methodPrefix)
        }
        if ($patch.methodPostfix) {
            $postfix = (New-Object HarmonyLib.HarmonyMethod -ArgumentList $patch.methodPostfix)
        }

        $patchResult = $harmony.Patch($patch.methodOrig, $prefix, $postfix)
        if (-not $patchResult) {
            throw "Could not patch $($patch.origName)"
        }
        else {
            Write-Host "Patched $($patch.origName)" -ForegroundColor Green
        }
    }

    $Global:patched = $true

    if ([SpmtModifications]::SkipListTemplateTypeCompatibilityCheck)
    {
        Write-Host "SkipListTemplateTypeCompatibilityCheck is activated" -ForegroundColor Green
    } else {
        Write-Host "SkipListTemplateTypeCompatibilityCheck is disabled" -ForegroundColor Gray
    }

    if ([SpmtModifications]::SkipUpdatingExistingViews)
    {
        Write-Host "SkipUpdatingExistingViews is activated" -ForegroundColor Green
    } else {
        Write-Host "SkipUpdatingExistingViews is disabled" -ForegroundColor Gray
    }

    if ([SpmtModifications]::SkipViewMigration)
    {
        Write-Host "SkipViewMigration is activated" -ForegroundColor Green
    } else {
        Write-Host "SkipViewMigration is disabled" -ForegroundColor Gray
    }
}

LoadSpmtPowerShellModule
LoadHarmony
CreateCustomTypeForHarmony
PatchSpmt