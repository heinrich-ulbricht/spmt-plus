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
  
    $Source = @"
    public class SpmtModifications
    {
      // note: __result contains the result value of the called method; in a Harmony postfix method we get the chance to modify the result value
      // see here for how this works: https://harmony.pardeike.net/articles/intro.html#how-harmony-works
  
      public static long Version = $Version;
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

      private static bool _skipUpdatingViews = false;
      public static bool SkipUpdatingViews {
        get {
            return _skipUpdatingViews;
        }
        set {
          System.Console.WriteLine("[HEU] Setting SkipUpdatingViews to " + value.ToString());
          Log("[HEU] Setting SkipUpdatingViews to " + value.ToString());
          _skipUpdatingViews = value;
        }
      }

      private static bool _skipViews = false;
      public static bool SkipViews {
        get {
            return _skipViews;
        }
        set {
          System.Console.WriteLine("[HEU] Setting SkipViews to " + value.ToString());
          Log("[HEU] Setting SkipViews to " + value.ToString());
          _skipViews = value;
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

      // note: this affects only some special views that already exist at the destination and are deemed internal; it is not called if a view does not yet exists at the destination
      public static void NeedSkipInternalViewPostfix(Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ViewMetaInfo view, Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo sourceList, ref bool __result)
      {
        string viewName = view.RelativePathToList.Trim('/');
        Log("[HEU] Original NeedSkipInternalView result for view '" + viewName + "' on list '" + sourceList.Title + "': " + __result.ToString());

        // note: ActivatePatches is not checked here; skipping is controlled by SkipUpdatingViews instead
        if (SkipUpdatingViews)
        {
            Log("[HEU] Skipping view '" + viewName + "' for list '" + sourceList.Title + "'");
            // skip all views
            __result = true;
        }
      }

      // this is called for every view of the source list
      public static bool UpsertViewPrefix(
        ref Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ViewMetaInfo __result,
        Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.ViewMigrator __instance,
        Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ViewMetaInfo view,
        Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo sourceList,
        Microsoft.SharePoint.MigrationTool.MigrationLib.Common.ListMetaInfo targetLibraryInfo
        )
      {
        if (!SkipViews)
        {
            return true;
        }
        Log("[HEU] Skipping UpsertView for '" + view.Title + "' for source list '" + sourceList.Url + "'");
        Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaObjectResult result = new Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaObjectResult()
          {
            ContentType = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationContentType.View,
            ContainerType = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationContentType.List,
            Title = view.Title,
            SourceId = view.Id,
            SourceUrl = sourceList.Url,
            targetUrl = targetLibraryInfo.Url
          };
          
          result.Operation = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationOperation.Skip;
          // this logs as "Skip the internal view" in the structure report
          result.Message = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationMessage.ViewSkipInternal;

          // result.Result = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationResult.Succeed; - don't (seems only to be set for non-skipped views)
          __instance.ReportMigrationResult(result);
          __result = view;

          return false;
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

    if ([SpmtModifications]::ActivatePatches)
    {
        Write-Host "Patches are activated" -ForegroundColor Green
    } else {
        Write-Host "Patches are disabled" -ForegroundColor Gray
    }

    if ([SpmtModifications]::SkipUpdatingViews)
    {
        Write-Host "Skipping updating views is activated" -ForegroundColor Green
    } else {
        Write-Host "Skipping updating views is disabled" -ForegroundColor Gray
    }

    if ([SpmtModifications]::SkipView)
    {
        Write-Host "Skipping views is activated" -ForegroundColor Green
    } else {
        Write-Host "Skipping views is disabled" -ForegroundColor Gray
    }
}
  
LoadSpmtPowerShellModule
LoadHarmony
CreateCustomTypeForHarmony
PatchSpmt