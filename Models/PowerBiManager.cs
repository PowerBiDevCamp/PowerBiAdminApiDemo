using System;
using System.Collections.Generic;
using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using System.IO;
using System.Net;
using System.Configuration;
using Newtonsoft.Json;
using System.Text;

namespace PowerBiAdminApiDemo.Models {

  class PowerBiManager {

    private readonly static string[] requiredScopes = PowerBiPermissionScopes.TenantReadWriteAll;

    static PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);
    //static PowerBIClient pbiClient = TokenManager.GetPowerBiAppOnlyClient();

    #region "User APIs (aka Non-admin APIs)

    public static void GetWorkspaces() {
      var workspaces = pbiClient.Groups.GetGroups().Value;
      foreach (var workspace in workspaces) {
        Console.WriteLine(workspace.Name);
      }
    }

    public static Group GetWorkspace(string WorkspaceName) {
      var workspaces = pbiClient.Groups.GetGroups().Value;
      foreach (var workspace in workspaces) {
        if (workspace.Name.Equals(WorkspaceName))
          return workspace;
      }
      return null;
    }

    public static void GetDatasetInWorkspace(Guid WorkspaceId, string DatasetId) {
      Dataset dataset = pbiClient.Datasets.GetDatasetInGroup(WorkspaceId, DatasetId);

    }

    public static Dataset GetDataset(Guid WorkspaceId, string DatasetName) {
      var datasets = pbiClient.Datasets.GetDatasetsInGroup(WorkspaceId).Value;
      foreach (var dataset in datasets) {
        if (dataset.Name.Equals(DatasetName)) {
          return dataset;
        }
      }
      return null;
    }

    public static IList<Dataset> GetDatasets(Guid WorkspaceId) {
      return pbiClient.Datasets.GetDatasetsInGroup(WorkspaceId).Value;
    }

    public static Report GetReport(Guid WorkspaceId, string ReportName) {
      var reports = pbiClient.Reports.GetReportsInGroup(WorkspaceId).Value;
      foreach (var report in reports) {
        if (report.Name.Equals(ReportName)) {
          return report;
        }
      }
      return null;
    }

    #endregion"

    #region "Admin APIs"

    public static void GetWorkspacesAsAdmin() {

      string workspaceFilter = "state eq 'Active' and type eq 'Workspace'";

      var workspaces = pbiClient.Groups.GetGroupsAsAdmin(top: 100, filter: workspaceFilter).Value;

      foreach (var workspace in workspaces) {
        // Console.WriteLine(workspace.Name);
      }

      SaveObjectAsJsonFile("Get-Workspaces-As-Admin.json", workspaces);
    }

    public static void GetWorkspacesAsAdminWithExpand() {

      string workspaceFilter = "state eq 'Active' and type eq 'Workspace'";
      string workspaceExpand = "users, reports, dashboards, datasets, dataflows";

      var workspaces = pbiClient.Groups.GetGroupsAsAdmin(top: 100,
                                                         filter: workspaceFilter,
                                                         expand: workspaceExpand).Value;

      SaveObjectAsJsonFile("Get-Workspaces-As-Admin-With-Expand.json", workspaces);
    }

    public static void GetWorkspaceAsAdmin(Guid WorkspaceId) {
      var workspace = pbiClient.Groups.GetGroupAsAdmin(WorkspaceId);
      SaveObjectAsJsonFile("Get-Workspace-As-Admin.json", workspace);
    }

    public static void GetWorkspaceUsersAsAdmin(Guid WorkspaceId) {
      var users = pbiClient.Groups.GetGroupUsersAsAdmin(WorkspaceId).Value;
      SaveObjectAsJsonFile("Get-Workspace-Users-As-Admin.json", users);
    }

    public static void AddWorkspaceUserAsAdmin(Guid WorkspaceId, string UserEmail) {

      // this operation must run as a user - service principal has read-only capabilities
      var pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var user = new GroupUser {
        PrincipalType = PrincipalType.User,
        EmailAddress = UserEmail,
        GroupUserAccessRight = GroupUserAccessRight.Admin
      };

      pbiClient.Groups.AddUserAsAdmin(WorkspaceId, user);
    }

    public static void AddWorkspaceServicePrincipalAsAdmin(Guid WorkspaceId, string ServicePrincpalId) {

      // this operation must run as a user - service principal has read-only capabilities
      var pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var user = new GroupUser {
        PrincipalType = PrincipalType.App,
        Identifier = ServicePrincpalId,
        GroupUserAccessRight = GroupUserAccessRight.Admin
      };

      pbiClient.Groups.AddUserAsAdmin(WorkspaceId, user);
    }

    public static void DeleteWorkspaceUserAsAdmin(Guid WorkspaceId, string Identifer) {

      // this operation must run as a user - service principal has read-only capabilities
      var pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      pbiClient.Groups.DeleteUserAsAdmin(WorkspaceId, Identifer);
    }

    public static void GetWorkspaceAsAdminWithExpand(Guid WorkspaceId) {
      var workspace = pbiClient.Groups.GetGroupAsAdmin(WorkspaceId, expand: "users, reports, dashboards, datasets, dataflows, workbooks");
      SaveObjectAsJsonFile("Get-Workspace-As-Admin-With-Expand.json", workspace);
    }

    public static void GetWorkspaceAsAdminByName(string WorkspaceName) {
      string workspaceFilter = "state eq 'Active' and type eq 'Workspace' and name eq '" + WorkspaceName + "'";
      var workspace = pbiClient.Groups.GetGroupsAsAdmin(top: 100, filter: workspaceFilter, expand: "reports").Value;
      SaveObjectAsJsonFile("Get-Workspace-As-Admin.json", workspace);
    }

    public static void GetCapacitiesAsAdmin() {

      var capacities = pbiClient.Admin.GetCapacitiesAsAdmin();

      SaveObjectAsJsonFile("Get-Capacities-As-Admin.json", capacities);
    }

    public static Capacity GetPrimaryCapacity() {
      var capacities = pbiClient.Admin.GetCapacitiesAsAdmin().Value;
      foreach (var capacity in capacities) {
        if (capacity.Sku.Equals("P1") && capacity.State.Equals("Active")) {
          return capacity;
        }
      }
      return null;
    }

    public static void GetCapacityUsersAsAdmin() {

      var capacity = GetPrimaryCapacity();
      var capacityUsers = pbiClient.Capacities.GetCapacityUsersAsAdmin(capacity.Id).Value;

      SaveObjectAsJsonFile("Get-Capacity-Users-As-Admin.json", capacityUsers);
    }

    public static void GetDatasetsInGroupAsAdmin(string WorkspaceName) {

      var workspace = GetWorkspace(WorkspaceName);
      var datasets = pbiClient.Datasets.GetDatasetsInGroupAsAdmin(workspace.Id).Value;

      SaveObjectAsJsonFile("Get-Datasets-In-Group-As-Admin.json", datasets);
    }

    public static void GetDatasetsAsAdmin() {
      var datasets = pbiClient.Datasets.GetDatasetsAsAdmin();
      SaveObjectAsJsonFile("Get-Datasets-As-Admin.json", datasets);
    }

    public static void GetDatasourcesAsAdmin(string WorkspaceName, string DatasetName) {
      var workspace = GetWorkspace(WorkspaceName);
      var dataset = GetDataset(workspace.Id, DatasetName);

      var datasources = pbiClient.Datasets.GetDatasourcesAsAdmin(dataset.Id);

      SaveObjectAsJsonFile("Get-Datasources-As-Admin.json", datasources);
    }

    public static void GetImportsAsAdmin() {

      var imports = pbiClient.Imports.GetImportsAsAdmin();

      SaveObjectAsJsonFile("Get-Imports-As-Admin.json", imports);
    }

    public static void GetRefreshablesAsAdmin() {
      var refreshables = pbiClient.Admin.GetRefreshables(100);
      SaveObjectAsJsonFile("Get-Refreshables-As-Admin.json", refreshables);
    }

    public static void GetReportsAsAdmin() {
      var reports = pbiClient.Reports.GetReportsAsAdmin(top: 1000);
      SaveObjectAsJsonFile("Get-Reports-As-Admin.json", reports);
    }

    public static void GetDashbardsAsAdmin() {
      var dashboards = pbiClient.Dashboards.GetDashboardsAsAdmin(top: 100);
      SaveObjectAsJsonFile("Get-Dashboards-As-Admin.json", dashboards);
    }

    public static void GetAppsAsAdmin() {
      var apps = pbiClient.Apps.GetAppsAsAdmin(top: 100);
      SaveObjectAsJsonFile("Get-Apps-As-Admin.json", apps);
    }

    public static void GetPipelinesAsAdmin() {
      var pipelines = pbiClient.Pipelines.GetPipelinesAsAdmin();
      SaveObjectAsJsonFile("Get-Pipelines-As-Admin.json", pipelines);
    }

    public static void GetDatasetUsersAsAdmin(string WorkspaceName, string DatasetName) {

      var workspace = GetWorkspace(WorkspaceName);
      var dataset = GetDataset(workspace.Id, DatasetName);
      var datasetId = new Guid(dataset.Id);
      var datasetUsers = pbiClient.Datasets.GetDatasetUsersAsAdmin(datasetId).Value;

      SaveObjectAsJsonFile("Get-Dataset-Users-As-Admin.json", datasetUsers);
    }

    public static void GetReportUsersAsAdmin(string WorkspaceName, string ReportName) {

      var workspace = GetWorkspace(WorkspaceName);
      var report = GetReport(workspace.Id, ReportName);

      var reportUsers = pbiClient.Reports.GetReportUsersAsAdmin(report.Id).Value;

      SaveObjectAsJsonFile("Get-Report-Users-As-Admin.json", reportUsers);
    }

    public static void GetAppUsersAsAdmin() {

      var apps = pbiClient.Apps.GetAppsAsAdmin(top: 100).Value;

      Dictionary<string, IList<AppUser>> appList = new Dictionary<string, IList<AppUser>>();

      foreach (var app in apps) {
        var appUsers = pbiClient.Apps.GetAppUsersAsAdmin(app.Id).Value;
        appList.Add(app.Name, appUsers);

      }

      SaveObjectAsJsonFile("Get-App-Users-As-Admin.json", appList);
    }

    public static void GetUserArtifactAccessAsAdmin(string UserId) {

      List<ArtifactAccessEntry> listArtifactAccessEntries = new List<ArtifactAccessEntry>();

      // execute first call - no continuation token required
      ArtifactAccessResponse artifactAccessResponse = pbiClient.Users.GetUserArtifactAccessAsAdmin(UserId);

      // add first set of artifacts into listArtifactAccessEntries collection variable
      listArtifactAccessEntries.AddRange(artifactAccessResponse.ArtifactAccessEntities);

      // continue making additon calls until ContinuationToken is null
      while (artifactAccessResponse.ContinuationToken != null) {

        // decode continuation token to use in next outbound call to GetUserArtifactAccessAsAdmin
        string formattedContinuationToken = $"'{WebUtility.UrlDecode(artifactAccessResponse.ContinuationToken)}'";

        // execute next call using continuation token 
        artifactAccessResponse = pbiClient.Users.GetUserArtifactAccessAsAdmin(UserId, formattedContinuationToken);

        // add next set of artifacts into listArtifactAccessEntries collection variable
        listArtifactAccessEntries.AddRange(artifactAccessResponse.ArtifactAccessEntities);

      } // fall out of while loop when request results does not contain a continuation token


      // get user name to parse into file name
      string userName = UserId.Substring(0, UserId.IndexOf("@"));

      SaveObjectAsJsonFile("Get-User-Artifact-Access-As-Admin-For-" + userName + ".json", listArtifactAccessEntries.ToArray());
    }

    public static void GetUserSubscriptionsAsAdmin(string UserId) {

      List<Subscription> listSubscriptions = new List<Subscription>();

      SubscriptionsByUserResponse subscriptionByUserResponse = pbiClient.Users.GetUserSubscriptionsAsAdmin(UserId);
      listSubscriptions.AddRange(subscriptionByUserResponse.SubscriptionEntities);

      if (subscriptionByUserResponse.ContinuationToken != null) {
        string formattedContinuationToken = $"'{WebUtility.UrlDecode(subscriptionByUserResponse.ContinuationToken)}'";
        subscriptionByUserResponse = pbiClient.Users.GetUserSubscriptionsAsAdmin(UserId, formattedContinuationToken);
        listSubscriptions.AddRange(subscriptionByUserResponse.SubscriptionEntities);
      }

      SaveObjectAsJsonFile("Get-User-Subscriptions-As-Admin.json", listSubscriptions.ToArray());
    }

    public static void GetUnusedArftifactsAsAdmin(string WorkspaceName) {
      var workspace = GetWorkspace(WorkspaceName);

      var artifactsResponse = pbiClient.Groups.GetUnusedArtifactsAsAdmin(workspace.Id);

      List<UnusedArtifactEntity> listUnusedArtifacts = new List<UnusedArtifactEntity>();
      listUnusedArtifacts.AddRange(artifactsResponse.UnusedArtifactEntities);

      while (artifactsResponse.ContinuationToken != null) {
        string formattedContinuationToken = $"'{WebUtility.UrlDecode(artifactsResponse.ContinuationToken)}'";
        artifactsResponse = pbiClient.Groups.GetUnusedArtifactsAsAdmin(workspace.Id, formattedContinuationToken);
        listUnusedArtifacts.AddRange(artifactsResponse.UnusedArtifactEntities);
      }

      SaveObjectAsJsonFile("Get-Unused-Arftifacts-As-Admin.json", listUnusedArtifacts);
    }

    public static void ScanWorkspaceAsAdmin(string WorkspaceName) {

      // get Id of target workspace
      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces);

      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    public static void ScanWorkspaceAsAdmin1(string WorkspaceName) {

      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      Group workspace = GetWorkspace(WorkspaceName);

      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces();
      requiredWorkspaces.Workspaces = new List<Guid?>();
      requiredWorkspaces.Workspaces.Add(workspace.Id);

      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         lineage: true);

      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
        System.Threading.Thread.Sleep(1000);
        Console.WriteLine(scanStatus.Status);
      }

      WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);

      SaveObjectAsJsonFile("Scan-Workspace-As-Admin1.json", scanResult);

    }

    public static void ScanWorkspaceAsAdminGetUsers(string WorkspaceName) {

      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         getArtifactUsers: true);

      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin-Get-Users.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    public static void ScanWorkspaceAsAdminWithLineage(string WorkspaceName) {

      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         lineage: true);

      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin-With-Lineage.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    public static void ScanWorkspaceAsAdminWithDatasourceDetails(string WorkspaceName) {

      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         lineage: true,
                                                                         datasourceDetails: true);

      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin-With-Datasource-Details.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    public static void ScanWorkspaceAsAdminWithDatassetSchema(string WorkspaceName) {

      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         datasetSchema: true);

      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin-With-Dataset-Schema.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    public static void ScanWorkspaceAsAdminWithDatassetExpressions(string WorkspaceName) {

      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         datasetSchema: true,
                                                                         datasetExpressions: true);



      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin-With-Dataset-Expressions.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    public static void ScanWorkspaceAsAdminWithEverything(string WorkspaceName) {

      Group workspace = GetWorkspace(WorkspaceName);
      Guid workspaceId = workspace.Id;

      // create RequiredWorkspaces parameter object used to call PostWorkspaceInfo
      RequiredWorkspaces requiredWorkspaces = new RequiredWorkspaces {
        Workspaces = new List<Guid?>() { workspaceId }
      };

      // start asynchronous workspace scan job
      ScanRequest scanStatus = pbiClient.WorkspaceInfo.PostWorkspaceInfo(requiredWorkspaces,
                                                                         getArtifactUsers: true,
                                                                         lineage: true,
                                                                         datasourceDetails: true,
                                                                         datasetSchema: true,
                                                                         datasetExpressions: true);

      // get ID of asynchronous workspace scan job
      Guid scanId = scanStatus.Id.Value;

      while (scanStatus.Status.Equals("NotStarted") || scanStatus.Status.Equals("Running")) {
        // take a secord or two before polling for success
        System.Threading.Thread.Sleep(1000);
        // continue to call GetScanStatus until job has completed
        scanStatus = pbiClient.WorkspaceInfo.GetScanStatus(scanId);
      }

      // get results after succesful scan
      if (scanStatus.Status.Equals("Succeeded")) {
        WorkspaceInfoResponse scanResult = pbiClient.WorkspaceInfo.GetScanResult(scanId);
        SaveObjectAsJsonFile("Scan-Workspace-As-Admin-With-Everything.json", scanResult);
      }

      // handle error that occurred during workspace scan
      if (scanStatus.Status.Equals("Failed")) {
        Console.WriteLine("Workspace Scanning Error: " + scanStatus.Error);
      }

    }

    #endregion

    #region "Admin API support for extracting usge data from Power BI event log"

    private static List<ActivityEventEntity> activityEvents = new List<ActivityEventEntity>();

    public static void GetActivityEvents(DateTime date) {

      string dateString = date.ToString("yyyy-MM-dd");
      Console.Write("Getting Power BI activity events for " + dateString);


      string startDateTime = "'" + dateString + "T00:00:00'";
      string endDateTime = "'" + dateString + "T23:59:59'";

      PowerBIClient pbiClient = TokenManager.GetPowerBiAppOnlyClient();
      ActivityEventResponse response = pbiClient.Admin.GetActivityEvents(startDateTime, endDateTime);

      ProcessActivityResponse(response);

      while (response.ContinuationToken != null) {
        string formattedContinuationToken = $"'{WebUtility.UrlDecode(response.ContinuationToken)}'";
        response = pbiClient.Admin.GetActivityEvents(null, null, formattedContinuationToken, null);
        ProcessActivityResponse(response);
      }
      
      Console.WriteLine();
      Console.WriteLine("Export process has exported " + activityEvents.Count + " events.");

      SaveObjectAsJsonFile(@"EventActivityLog-" + dateString + ".json", activityEvents);
      Console.WriteLine();

    }

    private static void ProcessActivityResponse(ActivityEventResponse response) {

      Console.Write(".");

      foreach (var activityEventEntity in response.ActivityEventEntities) {
        string activityEventEntityJson = JsonConvert.SerializeObject(activityEventEntity);
        ActivityEventEntity activityEvent = JsonConvert.DeserializeObject<ActivityEventEntity>(activityEventEntityJson);
        activityEvents.Add(activityEvent);
      }
    }


    #endregion

    #region "Support for exporting Admin API response data as JSON files"

    private static string ExportFolderPath = ConfigurationManager.AppSettings["export-folder-path"];

    private static void SaveObjectAsJsonFile(string FileName, object targetObject) {
      Console.WriteLine("Generating output file " + FileName);

      Stream exportFileStream = File.Create(ExportFolderPath + FileName);
      StreamWriter writer = new StreamWriter(exportFileStream);

      JsonSerializerSettings settings = new JsonSerializerSettings {
        DefaultValueHandling = DefaultValueHandling.Ignore,
        Formatting = Formatting.Indented
      };

      writer.Write(JsonConvert.SerializeObject(targetObject, settings));
      writer.Flush();
      writer.Close();
      exportFileStream.Close();

      // uncomment next line if you want the JSON file opened in Notepad
      // System.Diagnostics.Process.Start("notepad", ExportFolderPath + FileName);

    }

    #endregion"

  }
}



