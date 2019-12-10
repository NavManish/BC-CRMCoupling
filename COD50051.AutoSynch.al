codeunit 50051 D365_AutoSynch
{

    //TableNo = "Job Queue Entry";
    trigger OnRun()
    Var
    begin
        AutoSynchCRMEntities();
    end;


    procedure AutoSynchCRMEntities()
    var
        CRMFullSyncRevLine: record "CRM Full Synch. Review Line";
    begin
        //CRMFullSyncRevLine.DeleteAll();
        CRMFullSyncRevLine.Generate();

        //CRMFullSyncRevLine.Reset();
        //CRMFullSyncRevLine.SetFilter(Name, '<>%1&<>%2&<>%3', 'SALESORDER-ORDER', 'POSTEDSALESINV-INV', 'POSTEDSALESLINE-INV');
        //CRMFullSyncRevLine.DeleteAll();

        //CRMFullSyncRevLine.Reset();
        //Message('%1', CRMFullSyncRevLine.Count());
        CRMFullSyncRevLine.Start();

        Commit();
        //CRMFullSyncRevLine.DeleteAll();
    end;
}