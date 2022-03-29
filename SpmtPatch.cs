using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Migration.Common;
using Microsoft.SharePoint.Migration.Common.Exceptions;
using Microsoft.SharePoint.MigrationTool.MigrationLib.Common;
using Microsoft.SharePoint.MigrationTool.MigrationLib.Log;
using Microsoft.SharePoint.MigrationTool.MigrationLib.Schema;
using Microsoft.SharePoint.MigrationTool.MigrationLib.SharePoint;
using Microsoft.SharePoint.MigrationTool.PowerShell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text.RegularExpressions;

public class SpmtModifications
{
    // note: __result contains the result value of the called method; in a Harmony postfix method we get the chance to modify the result value
    // see here for how this works: https://harmony.pardeike.net/articles/intro.html#how-harmony-works

    public static long Version = 2;
    private static bool _skipListTemplateTypeCompatibilityCheck = false;
    public static bool SkipListTemplateTypeCompatibilityCheck
    {
        get
        {
            return _skipListTemplateTypeCompatibilityCheck;
        }
        set
        {
            Console.WriteLine("[HEU] Setting patches activated to " + value.ToString());
            Log("[HEU] Setting patches activated to " + value.ToString());
            _skipListTemplateTypeCompatibilityCheck = value;
        }
    }

    private static bool _skipUpdatingExistingViews = false;
    public static bool SkipUpdatingExistingViews
    {
        get
        {
            return _skipUpdatingExistingViews;
        }
        set
        {
            Console.WriteLine("[HEU] Setting SkipUpdatingExistingViews to " + value.ToString());
            Log("[HEU] Setting SkipUpdatingExistingViews to " + value.ToString());
            _skipUpdatingExistingViews = value;
        }
    }

    private static bool _skipViewMigration = false;
    public static bool SkipViewMigration
    {
        get
        {
            return _skipViewMigration;
        }
        set
        {
            Console.WriteLine("[HEU] Setting SkipViewMigration to " + value.ToString());
            Log("[HEU] Setting SkipViewMigration to " + value.ToString());
            _skipViewMigration = value;
        }
    }

    private static bool _enableSharePointOnlineAsSource = false;
    public static bool EnableSharePointOnlineAsSource
    {
        get
        {
            return _enableSharePointOnlineAsSource;
        }
        set
        {
            Console.WriteLine("[HEU] Setting EnableSharePointOnlineAsSource to " + value.ToString());
            Log("[HEU] Setting EnableSharePointOnlineAsSource to " + value.ToString());
            _enableSharePointOnlineAsSource = value;
        }
    }

    public static bool LogToConsole = false;
    public static HashSet<int> AdditionalListTemplateTypesToSupport = new HashSet<int>();

    private class TenantInfo
    {
        public string TenantName;
        public string UserName;
        public bool IsTarget;
    }

    private static List<TenantInfo> Tenants = new List<TenantInfo>();
    public static void RegisterTenant(string tenantName, string userName, bool isTarget)
    {
        if (Tenants.Where(t => t.TenantName.Equals(tenantName, StringComparison.OrdinalIgnoreCase)).Any())
        {
            // ignore double registrations to facilitate multiple runs of the same script
            return;
        }
        var value = new TenantInfo()
        {
            TenantName = tenantName,
            UserName = userName,
            IsTarget = isTarget
        };
        Tenants.Add(value);
    }

    public static void ValidateListMetaInfoPostfix(
        IList list,
        string listCultureInvariantTitle,
        SchemaMigrationMessage __result)
    {
        if (!SkipListTemplateTypeCompatibilityCheck)
        {
            return;
        }
        Log("[HEU] Original ValidateListMetaInfo result for '" + listCultureInvariantTitle + "': " + __result.ToString());

        // sample of explicitly handling one library by title
        if (listCultureInvariantTitle == "Style Library")
        {
            Log("[HEU] Changing result for Style Library");
            __result = SchemaMigrationMessage.None;
        }

        // or just allow everything...
        __result = SchemaMigrationMessage.None;
    }

    // note: this affects only some special views that already exist at the destination and are deemed internal; it is not called if a view does not yet exists at the destination
    public static void NeedSkipInternalViewPostfix(
        ViewMetaInfo view,
        ListMetaInfo sourceList,
        ref bool __result)
    {
        if (!SkipUpdatingExistingViews)
        {
            return;
        }

        string viewName = view.RelativePathToList.Trim('/');
        Log("[HEU] Original NeedSkipInternalView result for view '" + viewName + "' on list '" + sourceList.Title + "': " + __result.ToString());

        // note: SkipListTemplateTypeCompatibilityCheck is not checked here; skipping is controlled by SkipUpdatingExistingViews instead
        Log("[HEU] Skipping view '" + viewName + "' for list '" + sourceList.Title + "'");
        // skip all views
        __result = true;
    }

    public static void CheckOnPremSiteAccessibilityPostfix(
        ref bool __result,
        string siteUrl,
        string username,
        SecureString password)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] CheckOnPremSiteAccessibilityPostfix called for {0}, setting to true", siteUrl);
        __result = true;
    }

    public static bool GetSPOnPremDocumentListPrefix(
        ref SPAuthentication spAuth,
        string site)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return true;
        }

        Log("[HEU] GetSPOnPremDocumentList called for site '{0}', setting spAuth to null to force SPO source", site);
        spAuth = null;
        return true; // continue with original
    }

    public static void CreateClientContextPostfix(
        UserContext userContext,
        SPAuthentication spAuth,
        string siteUrl,
        string userAgent, IMigSPClientContext __result)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        // to aid debugging; the type must contain "Online" for SPO
        Log("[HEU] CreateClientContextPostfix for site '{0}' returns type {1}", siteUrl, __result.GetType().ToString());
    }

    private static UserContext Heu_CreateUserContextForSite(
        string siteUrl)
    {
        NetworkCredential sourceCred;
        if (SourceCredentials.TryGetValue(siteUrl, out sourceCred))
        {
            Log("[HEU] Creating userContext, getting credentials from cache for site '{0}', user is {1}", siteUrl, sourceCred.UserName);
            return new UserContext(sourceCred.UserName, sourceCred.Password);
        }
        else
        {
            Log("[HEU] Creating userContext, did not find credentials for site '{0}'; using dummy credential", siteUrl);
            return new UserContext("", "");
        }
    }

    public static void CreateClientContextPrefix(
        ref UserContext userContext,
        ref SPAuthentication spAuth,
        string siteUrl,
        string userAgent)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] CreateClientContextPrefix for site '{0}'", siteUrl);

        // necessary for SpPackageCreator
        if (spAuth != null)
        {
            Log("[HEU] CreateClientContextPrefix: setting spAuth to null for site {0}", siteUrl);
        }
        if (userContext == null)
        {
            Log("[HEU] CreateClientContextPrefix: creating userContext for site '{0}'", siteUrl);
            userContext = Heu_CreateUserContextForSite(siteUrl);
        }
    }

    public static void CanCreateOnPremClientContextPostfix(ref bool __result, SPAuthentication spAuth)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] CanCreateOnPremClientContext: would return '{0}'; forcing 'false'", __result);

        __result = false;
    }

    public static void InitAuthorizationHeaderPostfix(
        UserContext userContext)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] InitAuthorizationHeaderPostfix; userContext.CredentialUserName={0}, userContext.SpoUser={1}, userContext.TenantId={2}, userContext.AuthType={3}", userContext.CredentialUserName, userContext.SpoUser, userContext.TenantId, userContext.AuthType);
    }

    public static void SPODocumentAcquirer_CreateContextPrefix(
        SPAuthentication spAuth,
        string site,
        ref UserContext userContext)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        if (spAuth == null)
        {
            Log("[HEU] SPODocumentAcquirer_CreateContextPrefix called for " + site + " with spAuth==null");
            if (userContext == null)
            {
                Log("[HEU] Creating user context for " + site);

                userContext = Heu_CreateUserContextForSite(site);

            }
        }
        else
        {
            Log("[HEU] SPODocumentAcquirer_CreateContextPrefix called for " + site + " with spAuth set");
        }
    }

    public static void SharePointContext_CreateContextPrefix(
        SharePointContext __instance)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] SharePointContext_CreateContextPrefix: setting SharePointAuthentication to null");

        Log("[HEU] SharePointContext_CreateContextPrefix: GetCredentialsUser() " + __instance.GetCredentialsUser());
        Log("[HEU] SharePointContext_CreateContextPrefix: GetUserContext().SpoUser " + __instance.GetUserContext().SpoUser);
        Log("[HEU] SharePointContext_CreateContextPrefix: GetSiteUri() " + __instance.GetSiteUri());

        __instance.SharePointAuthentication = null;
    }

    public static void IsSourceVersionSupportedPostfix(
        ref bool __result,
        int version)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] Setting IsSourceVersionSupported for version " + version.ToString() + " to true");
        __result = true;
    }

    public static void IsSourceVersionSupported2Postfix(
        ref bool __result,
        IMigSPEnvironment migSpEnvironment)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] Setting IsSourceVersionSupported for version {0} to true", migSpEnvironment.SPVersion.ToString());
        __result = true;
    }

    public static void AcquireAccessTokenPostfix(
        IADAuthToken __result,
        string resourceURI,
        string username = null,
        SecureString password = null,
        bool isExternalProvidedURI = true)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }
        Log("[HEU] AcquireAccessTokenPostfix; username=" + username);
        Log("[HEU] Token.TenantId=" + __result.TenantId);
        Log("[HEU] Token.UserId=" + __result.UserId);
        //Log("[HEU] Token.AccessToken=" + __result.AccessToken);
        //Log("[HEU] Token.AuthorizationHeader=" + __result.AuthorizationHeader);
    }


    private static string Heu_GetResource(string resourceURI, bool isExternalProvidedURI)
    {
        UriBuilder uriBuilder;
        try
        {
            uriBuilder = new UriBuilder(resourceURI);
        }
        catch
        {
            throw new ArgumentException("[Auth] Illegal resource URI - " + resourceURI, "resourceURI");
        }
        if (!isExternalProvidedURI)
        {
            if (uriBuilder.Scheme != Uri.UriSchemeHttps)
                throw new ArgumentException("[Auth] resource URI must be HTTPS - " + resourceURI, "resourceURI");
            return resourceURI;
        }
        return new UriBuilder()
        {
            Host = uriBuilder.Host,
            Port = (uriBuilder.Uri.IsDefaultPort ? -1 : uriBuilder.Port),
            Scheme = Uri.UriSchemeHttps
        }.ToString();
    }

    public static bool AcquireAccessTokenPrefix(
        AADContext __instance,
        ref object ____objLock,
        ref string ___tokenCacheFileName,
        ref string ___tokenCacheDirectory,
        ref IPublicClientApplication ___app,
        ref IADAuthToken __result,
        string resourceURI,
        string username = null,
        SecureString password = null,
        bool isExternalProvidedURI = true)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return true;
        }
        //var stackTrace = Environment.StackTrace;
        var stackTrace = "logging disabled";
        Log("[HEU] AcquireAccessTokenPrefix; resourceURI={0}, username={1}, call stack={2}", resourceURI, username, stackTrace);
        if (string.IsNullOrEmpty(username))
        {
            // this is always the target as there was no graph for on-prem source site
            if (resourceURI.Contains("graph.microsoft.com"))
            {
                var targetTenant = Tenants.Where(t => t.IsTarget).First();
                username = targetTenant.UserName;
            }
            else
            {
                foreach (var tenant in Tenants)
                {
                    if (resourceURI.Contains(tenant.TenantName))
                    {
                        username = tenant.UserName;
                        break;
                    }
                }
            }
        }

        string[] scopes = new string[1]
        {
            string.Format("{0}/.default", Heu_GetResource(resourceURI, isExternalProvidedURI))
        };
        try
        {
            lock (____objLock)
            {
                if (___app == null)
                {
                    PublicClientApplicationOptions options = new PublicClientApplicationOptions();
                    options.ClientId = "fdd7719f-d61e-4592-b501-793734eb8a0e"; // this is the "new" ID; there is also an old ID in the code which we ignore here; same goes for redirect URI
                    options.LegacyCacheCompatibilityEnabled = true;
                    ___app = PublicClientApplicationBuilder.CreateWithApplicationOptions(options).WithAuthority(Office365Endpoints.GetEndpointURI(EndPointType.AUTHORITY)).WithRedirectUri(new Uri("https://login.microsoftonline.com/common/oauth2/nativeclient").ToString()).Build();
                    if (!string.IsNullOrEmpty(___tokenCacheFileName) && !string.IsNullOrEmpty(___tokenCacheDirectory))
                    {
                        LogManager.PreferedLogger.LogInformation(string.Format("TokenCacheDirectory: {0}, TokenCacheFileName: {1}", ___tokenCacheDirectory, ___tokenCacheFileName));
                        if (System.IO.File.Exists(Path.Combine(___tokenCacheDirectory, ___tokenCacheFileName)))
                            LogManager.PreferedLogger.LogInformation("TokenCache file exists and will be loaded.");
                        else
                            LogManager.PreferedLogger.LogInformation("TokenCache file does not exist. New file will be created after login.");
                        MsalCacheHelper.CreateAsync(new StorageCreationPropertiesBuilder(___tokenCacheFileName, ___tokenCacheDirectory).Build()).Result.RegisterCache(___app.UserTokenCache);
                    }
                }
            }
            IEnumerable<IAccount> result = ___app.GetAccountsAsync().Result;
            Log("[HEU] Got {0} authenticated accounts: {1}", result.Count(), string.Join(", ", result.Select(a => a.Username)));
            var account = result.Where(a => a.Username.Equals(username, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (null != account)
            {
                Log("[HEU] Found authenticated account for user name {0}", username);
                try
                {
                    __result = new AADAuthToken(___app.AcquireTokenSilent(scopes, account).ExecuteAsync().Result);
                    return false;
                }
                catch (MsalUiRequiredException)
                {
                    LogManager.PreferedLogger.LogDebug("MSAL require credential");
                }
            }
            else
            {
                Log("[HEU] Did not find account for user name {0}, need new authentication", username);
            }
            if (!string.IsNullOrEmpty(username) && password != null)
            {
                LogManager.PreferedLogger.LogDebug("Acquire token by username password");
                __result = new AADAuthToken(___app.AcquireTokenByUsernamePassword(scopes, username, password).ExecuteAsync().Result);
                return false;
            }
            LogManager.PreferedLogger.LogDebug("Acquire token by interaction");
            Console.WriteLine("Login for resource {0} and username {1}", resourceURI, username);
            __result = new AADAuthToken(___app.AcquireTokenInteractive(scopes).ExecuteAsync().Result);
            return false;
        }
        catch (AggregateException ex)
        {
            LogManager.PreferedLogger.LogException(ex, "Error happened in AAD login");
            throw new AADAuthenticationException(ex.InnerException);
        }
    }

    // this is called for every view of the source list
    public static bool UpsertViewPrefix(
        ref ViewMetaInfo __result,
        ViewMigrator __instance,
        ViewMetaInfo view,
        ListMetaInfo sourceList,
        ListMetaInfo targetLibraryInfo)
    {
        if (!SkipViewMigration)
        {
            return true;
        }
        Log("[HEU] Skipping UpsertView for '" + view.Title + "' for source list '" + sourceList.Url + "'");
        SchemaObjectResult result = new SchemaObjectResult()
        {
            ContentType = SchemaMigrationContentType.View,
            ContainerType = SchemaMigrationContentType.List,
            Title = view.Title,
            SourceId = view.Id,
            SourceUrl = sourceList.Url,
            targetUrl = targetLibraryInfo.Url
        };

        result.Operation = SchemaMigrationOperation.Skip;
        // this logs as "Skip the internal view" in the structure report
        result.Message = SchemaMigrationMessage.ViewSkipInternal;

        // result.Result = Microsoft.SharePoint.MigrationTool.MigrationLib.Schema.SchemaMigrationResult.Succeed; - don't (seems only to be set for non-skipped views)
        __instance.ReportMigrationResult(result);
        __result = view;

        return false;
    }

    public static void SharePointContext_ctorPrefix(
        ref object __instance,
        ref UserContext userContext,
        ref SPAuthentication spAuth,
        string siteUri,
        IMigSPEnvironment parentEnvironment = null,
        string userAgent = null)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }

        Log("[HEU] SharePointContext constructor: userContext={0}, spAuth={1}, siteUri={2}, parentEnvironment={3}, userAgent={4}", userContext, spAuth, siteUri, parentEnvironment, userAgent);
        spAuth = null;
        if (userContext == null)
        {
            // note: a valid user context does not seem to matter here...
            userContext = Heu_CreateUserContextForSite(siteUri);

        }
    }

    public static void AssessSpSiteAdminPrefix(
        ISharePointContext context)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }
        Log("[HEU] AssessSpSiteAdminPrefix for " + context.GetCredentialsUser());
    }

    public static void AssessSpSiteAdminPostfix(
        ISharePointContext context)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }
        Log("[HEU] AssessSpSiteAdminPrefix for " + context.GetCredentialsUser());
    }

    public static void IsFilteredListPostfix(
        IList docLib,
        ref bool __result)
    {
        if (!SkipListTemplateTypeCompatibilityCheck)
        {
            return;
        }
        Log("[HEU] Original IsFilteredList result for: '" + docLib.RootFolder.ServerRelativeUrl + "' (BaseTemplate " + docLib.BaseTemplate.ToString() + "): " + __result.ToString());

        // allow everything, filter nothing...
        __result = false;
    }

    public static bool CheckListTemplateMatchPrefix(
        ListMetaInfo sourceList,
        ListMetaInfo targetList,
        PackageSourceType SourceType)
    {
        if (!SkipListTemplateTypeCompatibilityCheck)
        {
            return true;
        }

        Log("[HEU] Skipping schema mismatch check");
        if (targetList.ListTemplateType != sourceList.ListTemplateType)
        {
            Log("[HEU] Using the target document library type {0} instead of the default {1} for list {2}", targetList.ListTemplateType, sourceList.ListTemplateType, targetList.Url);
            sourceList.ListTemplateType = targetList.ListTemplateType;
        }

        return false;
    }

    public static void PostUpdateListSchemaPrefix(
        ListMetaInfo sourceList,
        ListMetaInfo targetList)
    {
        if (!SkipListTemplateTypeCompatibilityCheck)
        {
            return;
        }
        Log("[HEU] PostUpdateListSchemaPrefix");

        Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract.ContentTypeManageDisable.Add(ListTemplateType.DocumentLibrary);
    }

    public static void PostUpdateListSchemaPostfix(
        ListMetaInfo sourceList,
        ListMetaInfo targetList)
    {
        if (!SkipListTemplateTypeCompatibilityCheck)
        {
            return;
        }
        Log("[HEU] PostUpdateListSchemaPostfix");

        Microsoft.SharePoint.MigrationTool.MigrationLib.Assessment.SchemaScanSpAbstract.ContentTypeManageDisable.Remove(ListTemplateType.DocumentLibrary);
    }

    // create instance of class with internal constructor
    public static T Heu_CreateInstance<T>(params object[] args)
    {
        var type = typeof(T);
        var instance = type.Assembly.CreateInstance(
            type.FullName, false,
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
            null, args, null, null);
        return (T)instance;
    }

    public static bool OpenBinaryDirectPrefix(
        ref FileInformation __result,
        IMigSPClientContext context,
        string serverRelativeUrl)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return true;
        }
        Log("[HEU] OpenBinaryDirectPrefix for " + serverRelativeUrl);
        // seems like we cannot use the provided "context" variable because the auth header there is missing; unclear why - so create a new context
        var clientContext = MIGUtilities.CreateClientContext(null, null, context.Url); ;
        var mStreamFromSpo = MIGUtilities.RunWithRetry(() =>
        {
            var file = clientContext.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
            clientContext.Load(file);
            Log("[HEU] OpenBinaryDirectPrefix: Executing query for " + serverRelativeUrl);

            clientContext.ExecuteQuery();
            IClientResult<Stream> streamResult = file.OpenBinaryStream();
            clientContext.ExecuteQuery();
            Log("[HEU] OpenBinaryDirectPrefix: Got stream for " + serverRelativeUrl);

            //var mStream = streamResult.Value;
            // note: there seem to be concurrency issues with pending requests when directly using the returned stream, so copy...
            MemoryStream mStream = new MemoryStream();
            streamResult.Value.CopyTo(mStream);
            mStream.Seek(0, SeekOrigin.Begin);
            Log("[HEU] OpenBinaryDirectPrefix: Memory stream length " + mStream.Length);
            return mStream;
        }, true, 3, 5, 20, "", typeof(WebExceptionBase));

        __result = Heu_CreateInstance<FileInformation>(mStreamFromSpo, Guid.NewGuid().ToString());
        while (context.HasPendingRequest)
        {
            context.ExecuteQuery();
        }
        return false;
    }

    public static void GetHttpResponsePrefix(
        ref string requestUrl,
        ISharePointContext context,
        [System.Runtime.CompilerServices.CallerMemberName] string callerMemberName = "",
        [System.Runtime.CompilerServices.CallerFilePath] string callerFilePath = "")
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }
        // need to change endpoints for getting versions; _vti_history does not support Bearer token, but newer REST endpoint does
        if (requestUrl.Contains("/_vti_history/"))
        {
            Log("[HEU] Changing /_vti_history/ URL " + requestUrl);

            var pattern = "(https://.+)(/(?:sites|personal)/.+/)(_vti_history/([0-9]+)/)";
            var s = requestUrl; // https://contoso.sharepoint.com/sites/2022-03-22-spmt-source/_vti_history/512/Freigegebene%20Dokumente/Requirements.docx
            var matches = Regex.Matches(s, pattern, RegexOptions.IgnoreCase);
            var hostUrl = matches[0].Groups[1].Value; // https://contoso.sharepoint.com
            var siteUrlPart = matches[0].Groups[2].Value; // /sites/2022-03-22-spmt-source/
            var historyPart = matches[0].Groups[3].Value; // _vti_history/512/
            var version = matches[0].Groups[4].Value; // 512
            var documentUri = new Uri(s.Replace(historyPart, "")); // https://contoso.sharepoint.com/sites/2022-03-22-spmt-source/Freigegebene Dokumente/Requirements.docx
            var serverRelativeUrl = documentUri.AbsolutePath;
            var restUrl = hostUrl + siteUrlPart + "_api/web/GetFileByServerRelativeUrl('" + serverRelativeUrl + "')/versions(" + version + ")/$value"; // https://contoso.sharepoint.com/sites/2022-03-22-spmt-source/_api/web/GetFileByServerRelativeUrl('/sites/2022-03-22-spmt-source/Freigegebene%20Dokumente/Requirements.docx')/versions(512)/$value
            requestUrl = restUrl;
        }
    }

    static Dictionary<string, NetworkCredential> SourceCredentials = new Dictionary<string, NetworkCredential>();
    public static void ProcessRecordPrefix(AddSPMTTask __instance)
    {
        if (!EnableSharePointOnlineAsSource)
        {
            return;
        }
        var cred = __instance.SharePointSourceCredential.GetNetworkCredential();
        Log("[HEU] Caching credential for user {0} for site {1}", __instance.SharePointSourceSiteUrl, cred.UserName);
        SourceCredentials[__instance.SharePointSourceSiteUrl] = cred;
    }

    public static void GetSupportedListTemplatesPostfix(ref HashSet<ListTemplateType> __result)
    {
        if (!SkipListTemplateTypeCompatibilityCheck)
        {
            return;
        }
        // note: this is the original SchemaScanSpAbstract.SupportedListTemplates plus added template types we want SPMT to support
        __result = new HashSet<ListTemplateType>()
        {
            ListTemplateType.DocumentLibrary,
            ListTemplateType.MySiteDocumentLibrary,
            ListTemplateType.PictureLibrary,
            ListTemplateType.Announcements,
            ListTemplateType.Contacts,
            ListTemplateType.GanttTasks,
            ListTemplateType.GenericList,
            ListTemplateType.DiscussionBoard,
            ListTemplateType.Events,
            ListTemplateType.IssueTracking,
            ListTemplateType.Links,
            ListTemplateType.Tasks,
            ListTemplateType.Survey,
            ListTemplateType.TasksWithTimelineAndHierarchy,
            ListTemplateType.XMLForm,
            ListTemplateType.Posts,
            ListTemplateType.Comments,
            ListTemplateType.Categories,
            ListTemplateType.CustomGrid,
            ListTemplateType.WebPageLibrary,
            ListTemplateType.PromotedLinks,
            (ListTemplateType) 500,
            (ListTemplateType) 851,
            (ListTemplateType) 880,
            (ListTemplateType) 1302,
            (ListTemplateType) 10015,
            (ListTemplateType) 10017,
            (ListTemplateType) 10018,
            (ListTemplateType) 10019,
            (ListTemplateType) 5001, // added: Nintex
            (ListTemplateType) 550, // added: Social
            (ListTemplateType) 544 // added: MicroFeed
            // add new types via AdditionalListTemplateTypesToSupport
        };
        foreach (var templateType in AdditionalListTemplateTypesToSupport)
        {
            __result.Add((ListTemplateType)templateType);
        }
    }

    public static void Log(string errorMessage = "", params object[] messageArgs)
    {
        if (LogToConsole && !errorMessage.Contains("[VER]") && !errorMessage.Contains("[DBG]"))
        {
            Console.WriteLine(string.Format(errorMessage, messageArgs));
        }

        bool retry;
        // solve concurrency issues by retrying; not pretty, but works
        do
        {
            try
            {
                System.IO.File.AppendAllText("spmtplus.log", string.Format(errorMessage + Environment.NewLine, messageArgs));
                retry = false;
            }
            catch
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

    public static void LogExceptionPrefix(Exception e, string errorMessage = "", params object[] messageArgs)
    {
        Log("[EXC] " + errorMessage + " (" + e.ToString() + ")", messageArgs);
    }

    public static IFolder Heu_EnsureFolderExistance(string targetWebUrl, string targetDocLib, string targetSubFolder)
    {
        var ctx = MIGUtilities.CreateClientContext(null, null, targetWebUrl);
        return MIGUtilities.RunWithRetry(() =>
        {
            var web = ctx.Web;
            // "Dokumente"
            var list = web.Lists.GetByTitle(targetDocLib);
            ctx.Load(list, l => l.RootFolder);
            ctx.ExecuteQuery();

            // "Freigegebene Dokumente"
            var rootFolderName = list.RootFolder.Name;
            // "/sites/test/Freigegebene Dokumente"
            var rootFolderUrl = list.RootFolder.ServerRelativeUrl;
            var targetSubFolderPathParts = targetSubFolder.Split('/');

            var parentFolder = list.RootFolder;
            var folderToCheckUrl = rootFolderUrl;
            foreach (var folderName in targetSubFolderPathParts)
            {
                folderToCheckUrl += "/" + folderName;
                Log("[HEU] Checking existance of folder {0}", folderToCheckUrl);
                var folder = web.GetFolderByServerRelativeUrl(folderToCheckUrl);
                ctx.Load(folder);
                try
                {
                    ctx.ExecuteQuery();
                    Log("[HEU] Folder exists: {0}", folderName);
                }
                catch (MigSPNotFoundException)
                {
                    Log("[HEU] Did not find it. Creating folder {0}", folderName);
                    parentFolder.Folders.Add(folderName);
                    ctx.ExecuteQuery();
                    folder = web.GetFolderByServerRelativeUrl(folderToCheckUrl);
                    ctx.ExecuteQuery();
                    Log("[HEU] Successfully created folder at {0}", folderToCheckUrl);
                }
                parentFolder = folder;
            }
            return list.RootFolder;

            //ctx.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl)
        }, true, 3, 5, 20, "", typeof(WebExceptionBase));
    }
}