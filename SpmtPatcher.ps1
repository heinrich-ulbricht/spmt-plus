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
    Set-Location $spmtModulePath
    $null = Import-Module "Microsoft.SharePoint.MigrationTool.PowerShell"
    Set-Location $currentLocation
  
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
    $Version = 1
  
    $spmtModificationsType = ([System.Management.Automation.PSTypeName]'SpmtModifications').Type
    # check if version matches or if we got an old type in memory
    if ($spmtModificationsType) {
        if ([SpmtModifications]::Version -ne $Version) {
            throw "Detected old SpmtModifications type in memory. This can happen after code modifications. Please restart the PowerShell process and run the script again."
        }
        Write-Host "Found existing custom SPMT modification type, good"
        return
    }
  
    $Source = @"
    public class SpmtModifications
    {
      // note: __result contains the result value of the called method; in a Harmony postfix method we get the chance to modify the result value
      // see here for how this works: https://harmony.pardeike.net/articles/intro.html#how-harmony-works
  
      public static long Version = 1;
      private static bool _activatePatches = true;
      public static bool ActivatePatches {
          get {
              return _activatePatches;
          }
          set {
            System.Console.WriteLine("[HEU] Setting patches activated to " + value.ToString());
            Log("[HEU] Setting patches activated to " + value.ToString());
              _activatePatches = value;
          }
      }

      public static bool LogToConsole = false;

      public static void ValidateListMetaInfoPostfix(Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.IList list, string listCultureInvariantTitle, ref Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationMessage __result)
      {
          Log("[HEU] Original ValidateListMetaInfo result for '" + listCultureInvariantTitle + "': " + __result.ToString());
  
          if (!ActivatePatches)
          {
              return;
          }

          // sample of explicitly handling one library by title
          if (listCultureInvariantTitle == "Style Library")
          {
            Log("[HEU] Changing result for Style Library");
            __result = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationMessage.None;
          }
    
          // or just allow everything...
          __result = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationMessage.None;
      }
    
      public static void IsFilteredListPostfix(Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.IList docLib, ref bool __result)
      {
          Log("[HEU] Original IsFilteredList result for: '" + docLib.RootFolder.ServerRelativeUrl + "' (BaseTemplate " + docLib.BaseTemplate.ToString() + "): " + __result.ToString());
  
          if (!ActivatePatches)
          {
              return;
          }
          // allow everything, filter nothing...
          __result = false;
      }
  
      public static bool CheckListTemplateMatchPrefix(Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo sourceList, Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo targetList, Microsoft.SharePoint.MigrationTool.MigrationLib.Common.PackageSourceType SourceType)
      {
          Log("[HEU] Skipping schema mismatch check");
          if (targetList.ListTemplateType != sourceList.ListTemplateType)
          {
            Log("[HEU] Using the target document library type {0} instead of the default {1} for list {2}", (object) targetList.ListTemplateType, (object) sourceList.ListTemplateType, (object) targetList.Url);
            sourceList.ListTemplateType = targetList.ListTemplateType;
          }
          
          if (!ActivatePatches)
          {
              return true;
          }

          return false;
      }
  
      public static void PostUpdateListSchemaPrefix(Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo sourceList, Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo targetList)
      {
          Log("[HEU] PostUpdateListSchemaPrefix");

          if (!ActivatePatches)
          {
              return;
          }

          Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract.ContentTypeManageDisable.Add(Microsoft.SharePoint.Client.ListTemplateType.DocumentLibrary);
      }
  
      public static void PostUpdateListSchemaPostfix(Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo sourceList, Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo targetList)
      {
        Log("[HEU] PostUpdateListSchemaPostfix");

        if (!ActivatePatches)
        {
            return;
        }

        Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract.ContentTypeManageDisable.Remove(Microsoft.SharePoint.Client.ListTemplateType.DocumentLibrary);
      }
  
      public static void Log(string errorMessage = "", params object[] messageArgs)
      {
        if (LogToConsole && !errorMessage.Contains("[VER]") && !errorMessage.Contains("[DBG]"))
        {
            System.Console.WriteLine(string.Format(errorMessage, messageArgs));
        }

        bool retry;
        // solve concurrency issues by retrying; not pretty, but works
        do
        {
          try
          {
            System.IO.File.AppendAllText("spmtplus.log", string.Format(errorMessage+System.Environment.NewLine, messageArgs));
            retry = false;
          } catch
          {
            retry = true;
          }
        } while (retry);
      }
  
      public static void LogDebugPrefix(string errorMessage = "", params object[] messageArgs)
      {
        Log("[DBG] " + errorMessage, messageArgs);
      }
  
      public static void LogInformationPrefix(string errorMessage = "", params object[] messageArgs)
      {
        Log("[INF] " + errorMessage, messageArgs);
      }
  
      public static void LogWarningPrefix(string errorMessage = "", params object[] messageArgs)
      {
        Log("[WRN] " + errorMessage, messageArgs);
      }
  
      public static void LogErrorPrefix(string errorMessage = "", params object[] messageArgs)
      {
        Log("[ERR] " + errorMessage, messageArgs);
      }
  
      public static void LogVerbosePrefix(string errorMessage = "", params object[] messageArgs)
      {
        Log("[VER] " + errorMessage, messageArgs);
      }
  
      public static void LogExceptionPrefix(System.Exception e, string errorMessage = "", params object[] messageArgs)
      {
        Log("[EXC] " + errorMessage + " (" + e.ToString() + ")", messageArgs);
      }
  
      public static void GetSupportedListTemplatesPostfix(ref System.Collections.Generic.HashSet<Microsoft.SharePoint.Client.ListTemplateType> __result)
      {
        if (!ActivatePatches)
        {
            return;
        }
        // note: this is the original SchemaScanSpAbstract.SupportedListTemplates plus added template types we want SPMT to support
        __result = new System.Collections.Generic.HashSet<Microsoft.SharePoint.Client.ListTemplateType>()
          {
            Microsoft.SharePoint.Client.ListTemplateType.DocumentLibrary,
            Microsoft.SharePoint.Client.ListTemplateType.MySiteDocumentLibrary,
            Microsoft.SharePoint.Client.ListTemplateType.PictureLibrary,
            Microsoft.SharePoint.Client.ListTemplateType.Announcements,
            Microsoft.SharePoint.Client.ListTemplateType.Contacts,
            Microsoft.SharePoint.Client.ListTemplateType.GanttTasks,
            Microsoft.SharePoint.Client.ListTemplateType.GenericList,
            Microsoft.SharePoint.Client.ListTemplateType.DiscussionBoard,
            Microsoft.SharePoint.Client.ListTemplateType.Events,
            Microsoft.SharePoint.Client.ListTemplateType.IssueTracking,
            Microsoft.SharePoint.Client.ListTemplateType.Links,
            Microsoft.SharePoint.Client.ListTemplateType.Tasks,
            Microsoft.SharePoint.Client.ListTemplateType.Survey,
            Microsoft.SharePoint.Client.ListTemplateType.TasksWithTimelineAndHierarchy,
            Microsoft.SharePoint.Client.ListTemplateType.XMLForm,
            Microsoft.SharePoint.Client.ListTemplateType.Posts,
            Microsoft.SharePoint.Client.ListTemplateType.Comments,
            Microsoft.SharePoint.Client.ListTemplateType.Categories,
            Microsoft.SharePoint.Client.ListTemplateType.CustomGrid,
            Microsoft.SharePoint.Client.ListTemplateType.WebPageLibrary,
            Microsoft.SharePoint.Client.ListTemplateType.PromotedLinks,
            (Microsoft.SharePoint.Client.ListTemplateType) 500,
            (Microsoft.SharePoint.Client.ListTemplateType) 851,
            (Microsoft.SharePoint.Client.ListTemplateType) 880,
            (Microsoft.SharePoint.Client.ListTemplateType) 1302,
            (Microsoft.SharePoint.Client.ListTemplateType) 10015,
            (Microsoft.SharePoint.Client.ListTemplateType) 10017,
            (Microsoft.SharePoint.Client.ListTemplateType) 10018,
            (Microsoft.SharePoint.Client.ListTemplateType) 10019,
            (Microsoft.SharePoint.Client.ListTemplateType) 5001, // added: Nintex
            (Microsoft.SharePoint.Client.ListTemplateType) 550, // added: Social
            (Microsoft.SharePoint.Client.ListTemplateType) 544 // added: MicroFeed
            // [EXTENSION POINT] add the types you need; don't forget the comma in the previous line
          }; 
      } 
    }
"@
  
    $spmtModule = Get-Module "Microsoft.SharePoint.MigrationTool.PowerShell"
    $spmtModulePath = (Get-Item $spmtModule.Path).Directory.FullName
  
    Write-Host "Loading types from SPMT module path '$spmtModulePath'"
    [void][reflection.assembly]::LoadFrom("$spmtModulePath\microsoft.sharepoint.migrationtool.migrationlib.dll")
    $null = Add-Type -TypeDefinition $Source -ReferencedAssemblies "mscorlib", "$spmtModulePath\microsoft.sharepoint.client.dll", "$spmtModulePath\microsoft.sharepoint.migrationtool.migrationlib.dll", "System.Console", "$spmtModulePath\microsoft.sharepoint.client.runtime.dll", "System.Core", "System.Runtime" -ErrorAction Continue
  
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
  
    $harmony = New-Object HarmonyLib.Harmony -ArgumentList "com.heinrich.spmtredirects"
    $validateListMetaInfoOrig = [Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract].GetMethod("ValidateListMetaInfo", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::NonPublic)
    $validateListMetaInfoPostfix = [SpmtModifications].GetMethod("ValidateListMetaInfoPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)  
    $patchResult = $harmony.Patch($validateListMetaInfoOrig, $null, (New-Object HarmonyLib.HarmonyMethod -ArgumentList $validateListMetaInfoPostfix))
    if (-not $patchResult) {
        throw "Could not patch ValidateListMetaInfo"
    }
    else {
        Write-Host "Patched ValidateListMetaInfo" -ForegroundColor Green
    }
    
    $isFilteredListOrig = [Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint.SPODocumentAcquirer].GetMethod("IsFilteredList", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::NonPublic)
    $isFilteredListPostfix = [SpmtModifications].GetMethod("IsFilteredListPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    $patchResult = $harmony.Patch($isFilteredListOrig, $null, (New-Object HarmonyLib.HarmonyMethod -ArgumentList $isFilteredListPostfix))
    if (-not $patchResult) {
        throw "Could not patch IsFilteredList"
    }
    else {
        Write-Host "Patched IsFilteredList" -ForegroundColor Green
    }
  
    $checkListTemplateMatchOrig = [Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaUtilities].GetMethod("CheckListTemplateMatch", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    $checkListTemplateMatchPrefix = [SpmtModifications].GetMethod("CheckListTemplateMatchPrefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    $patchResult = $harmony.Patch($checkListTemplateMatchOrig, (New-Object HarmonyLib.HarmonyMethod -ArgumentList $checkListTemplateMatchPrefix), $null)
    if (-not $patchResult) {
        throw "Could not patch CheckListTemplateMatch"
    }
    else {
        Write-Host "Patched CheckListTemplateMatch" -ForegroundColor Green
    }
  
    # remove types you don't want to log
    $logTypes = "Information", "Error", "Warning", "Exception", "Debug", "Verbose"
    foreach ($logType in $logTypes) {
        $logDebugOrig = [Microsoft.SharePoint.MigrationTool.MigrationLib.Common.MigObjectAbstract].GetMethod("Log$($logType)", [System.Reflection.BindingFlags]::Instance + [System.Reflection.BindingFlags]::Public)
        $logDebugPrefix = [SpmtModifications].GetMethod("Log$($logType)Prefix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
        $patchResult = $harmony.Patch($logDebugOrig, (New-Object HarmonyLib.HarmonyMethod -ArgumentList $logDebugPrefix), $null)
        if (-not $patchResult) {
            throw "Could not patch Log$($logType)"
        }
        else {
            Write-Host "Patched Log$($logType)" -ForegroundColor Green
        }
    }
  
    $prop = [Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract].GetProperty("SupportedListTemplates", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    $getter = $prop.GetAccessors()
    $getSupportedListTemplatesOrig = $getter[0]
    $getSupportedListTemplatesPostfix = [SpmtModifications].GetMethod("GetSupportedListTemplatesPostfix", [System.Reflection.BindingFlags]::Static + [System.Reflection.BindingFlags]::Public)
    $patchResult = $harmony.Patch($getSupportedListTemplatesOrig, $null, (New-Object HarmonyLib.HarmonyMethod -ArgumentList $getSupportedListTemplatesPostfix))  
    if (-not $patchResult) {
        throw "Could not patch GetSupportedListTemplates"
    }
    else {
        Write-Host "Patched GetSupportedListTemplates" -ForegroundColor Green
    }
  
    $Global:patched = $true

    if ([SpmtModifications]::ActivatePatches)
    {
        Write-Host "Patches are activated" -ForegroundColor Green
    } else {
        Write-Host "Patches are disabled" -ForegroundColor Gray
    }
}
  
LoadSpmtPowerShellModule
LoadHarmony
CreateCustomTypeForHarmony
PatchSpmt