using System;
using PowerBiAdminApiDemo.Models;

namespace PowerBiAdminApiDemo {
  class Program {

    const string WorkspaceName = "Contoso Sales Dev";
    const string WorkspaceName2 = "Multi-language Reports";
    const string ReportName = "Customer Sales";
    const string DatasetName = "Customer Sales";
    const string DatasetName2 = "ProductSales-Multi-Language-Demo";
    const string User1 = "tedp@powerbidevcamp.net";
    const string User2 = "austinp@powerbidevcamp.net";

    static void Main() {
      RunDemo01();
      //RunDemo02();
      //RunDemo03();
      //RunDemo04();
      //RunDemo05();
      //RunDemo06();
    }

    static void RunDemo01() {

      PowerBiManager.GetWorkspacesAsAdmin();
      
      PowerBiManager.GetWorkspacesAsAdminWithExpand();

      var workspace = PowerBiManager.GetWorkspace(WorkspaceName);
      PowerBiManager.GetWorkspaceAsAdmin(workspace.Id);

      PowerBiManager.GetWorkspaceAsAdminWithExpand(workspace.Id);

    }

    static void RunDemo02() {

      var workspace = PowerBiManager.GetWorkspace(WorkspaceName);

      // PowerBiManager.AddWorkspaceUserAsAdmin(workspace.Id, User2);
      
      PowerBiManager.GetWorkspaceUsersAsAdmin(workspace.Id);

      // PowerBiManager.DeleteWorkspaceUserAsAdmin(workspace.Id, "jackr@powerbidevcamp.net");
      
    }

    static void RunDemo03() {

      PowerBiManager.GetCapacitiesAsAdmin();

      PowerBiManager.GetCapacityUsersAsAdmin();
      
      PowerBiManager.GetDatasetsInGroupAsAdmin(WorkspaceName);
      
      PowerBiManager.GetDatasourcesAsAdmin(WorkspaceName2, DatasetName2);
      
      PowerBiManager.GetRefreshablesAsAdmin();
      
      PowerBiManager.GetReportsAsAdmin();
      
      PowerBiManager.GetDashbardsAsAdmin();
      
      PowerBiManager.GetImportsAsAdmin();
      
      PowerBiManager.GetAppsAsAdmin();

    }

    static void RunDemo04() {
    
      PowerBiManager.GetReportUsersAsAdmin(WorkspaceName, ReportName);
      
      PowerBiManager.GetDatasetUsersAsAdmin(WorkspaceName, DatasetName);
      
      PowerBiManager.GetAppUsersAsAdmin();
      
      PowerBiManager.GetUserArtifactAccessAsAdmin(User1);
      
      PowerBiManager.GetUserSubscriptionsAsAdmin(User1);
      
      PowerBiManager.GetUnusedArftifactsAsAdmin(WorkspaceName);

    }

    static void RunDemo05() {

      PowerBiManager.ScanWorkspaceAsAdmin(WorkspaceName);
      
      PowerBiManager.ScanWorkspaceAsAdminGetUsers(WorkspaceName);
      
      PowerBiManager.ScanWorkspaceAsAdminWithLineage(WorkspaceName);
      
      PowerBiManager.ScanWorkspaceAsAdminWithDatasourceDetails(WorkspaceName);
      
      PowerBiManager.ScanWorkspaceAsAdminWithDatassetSchema(WorkspaceName);
      
      PowerBiManager.ScanWorkspaceAsAdminWithDatassetExpressions(WorkspaceName);
    }

    static void RunDemo06() {

      DateTime date1 = new DateTime(2022, 2, 21);
      PowerBiManager.GetActivityEvents(date1);

      DateTime date2 = new DateTime(2022, 2, 18);
      PowerBiManager.GetActivityEvents(date2);

    }

  }
}
