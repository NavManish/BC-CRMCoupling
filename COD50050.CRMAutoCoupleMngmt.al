codeunit 50050 "D365_CRM AutoCouple Mngmt"
{

    trigger OnRun()
    begin
    end;

    var

        CRMProductName: Codeunit "CRM Product Name";
        RemoveCoupledContactsUnderCustomerQst: Label 'The Customer and %2 Account have %1 child Contact records coupled to one another. Do you want to delete their couplings as well?', Comment = '%1 is a number, %2 is CRM Product Name';


    procedure IsRecordCoupledToCRM(RecordID: RecordID): Boolean
    var
        CRMIntegrationRecord: Record "CRM Integration Record";
    begin
        exit(CRMIntegrationRecord.IsRecordCoupled(RecordID));
    end;

    procedure IsRecordCoupledToNAV(CRMID: Guid; NAVTableID: Integer): Boolean
    var
        CRMIntegrationRecord: Record "CRM Integration Record";
        NAVRecordID: RecordID;
    begin
        exit(CRMIntegrationRecord.FindRecordIDFromID(CRMID, NAVTableID, NAVRecordID));
    end;

    local procedure AssertTableIsMapped(TableID: Integer)
    var
        IntegrationTableMapping: Record "Integration Table Mapping";
    begin
        IntegrationTableMapping.SetRange("Table ID", TableID);
        IntegrationTableMapping.FindFirst();
        //IntegrationTableMapping.Get(TableID);
    end;

    procedure DefineCoupling(RecordID: RecordID; var CRMID: Guid; var CreateNew: Boolean; var Synchronize: Boolean; var Direction: Option): Boolean
    var
        CRMIntegrationRecord: Record "CRM Integration Record";
        CouplingRecordBuffer: Record "Coupling Record Buffer";
        CRMCouplingRecord: Page "CRM Coupling Record";
    begin
        AssertTableIsMapped(RecordID.TableNo());
        CRMCouplingRecord.SetSourceRecordID(RecordID);
        if CRMCouplingRecord.RunModal() = ACTION::OK then begin
            CRMCouplingRecord.GetRecord(CouplingRecordBuffer);
            if CouplingRecordBuffer."Create New" then
                CreateNew := true
            else
                if not IsNullGuid(CouplingRecordBuffer."CRM ID") then begin
                    CRMID := CouplingRecordBuffer."CRM ID";
                    CRMIntegrationRecord.CoupleRecordIdToCRMID(RecordID, CouplingRecordBuffer."CRM ID");
                    if CouplingRecordBuffer.GetPerformInitialSynchronization() then begin
                        Synchronize := true;
                        Direction := CouplingRecordBuffer.GetInitialSynchronizationDirection();
                    end;
                end else
                    exit(false);
            exit(true);
        end;
        exit(false);
    end;

    Local procedure SetAutoCoupling(RecordID: RecordID; var CRMID: Guid; var CreateNew: Boolean; var Synchronize: Boolean; var Direction: Option; CRMCouplingName: Text[250]): Boolean
    var
        CRMIntegrationRecord: Record "CRM Integration Record";
        CouplingRecordBuffer: Record "Coupling Record Buffer";
        CRMCoupledFields: Page "CRM Coupled Fields";
    begin

        if CreateNew then
            exit(true);

        AssertTableIsMapped(RecordID.TableNo());

        // --- START
        //CRMCouplingRecord.SetSourceRecordID(RecordID);
        CouplingRecordBuffer.Initialize(RecordID);
        IF NOT CouplingRecordBuffer.Insert() then begin
            CouplingRecordBuffer.Validate("CRM Name", CRMCouplingName);
            CouplingRecordBuffer."Sync Action" := CouplingRecordBuffer."Sync Action"::"To Integration Table";
            CouplingRecordBuffer.Modify();
        END ELSE BEGIN
            CouplingRecordBuffer."Sync Action" := CouplingRecordBuffer."Sync Action"::"To Integration Table";
            CouplingRecordBuffer.Validate("CRM Name", CRMCouplingName);
            CouplingRecordBuffer.Modify();
        END;
        CRMCoupledFields.SetSourceRecord(CouplingRecordBuffer);

        //CouplingRecordBuffer.Modify();
        //CRMCoupledFields.SetSourceRecord(CouplingRecordBuffer);
        // --- END
        // if CRMCouplingRecord.RunModal() = ACTION::OK then begin
        //     CRMCouplingRecord.GetRecord(CouplingRecordBuffer);
        //     if CouplingRecordBuffer."Create New" then
        //         CreateNew := true
        //    else
        if not IsNullGuid(CouplingRecordBuffer."CRM ID") then begin
            CRMID := CouplingRecordBuffer."CRM ID";
            CRMIntegrationRecord.CoupleRecordIdToCRMID(RecordID, CouplingRecordBuffer."CRM ID");
            if CouplingRecordBuffer.GetPerformInitialSynchronization() then begin
                Synchronize := true;
                Direction := CouplingRecordBuffer.GetInitialSynchronizationDirection();
            end;
        end else
            exit(false);
        exit(true);
        // end;
        // exit(false);
    end;

    procedure DefineAutoCoupling(RecordID: RecordID; CRMCouplingName: Text[250]; CreateNew: Boolean): Boolean
    var
        CRMIntegrationMngmtCU: Codeunit "CRM Integration Management";
        CRMID: Guid;
        Direction: Option;
        Synchronize: Boolean;
    begin
        RemoveCoupling(RecordID);
        Commit();
        SetAutoCoupling(RecordID, CRMID, CreateNew, Synchronize, Direction, CRMCouplingName);
        if CreateNew then
            CRMIntegrationMngmtCU.CreateNewRecordsInCRM(RecordID)
        else
            if Synchronize then
                CRMIntegrationMngmtCU.CoupleCRMEntity(RecordID, CRMID, Synchronize, Direction);
        //Synchronize := true;
    end;

    procedure RemoveCoupling(RecordID: RecordID)
    var
        TempCRMIntegrationRecord: Record "CRM Integration Record" temporary;
    begin
        RemoveCouplingWithTracking(RecordID, TempCRMIntegrationRecord);
    end;

    procedure RemoveCouplingWithTracking(RecordID: RecordID; var TempCRMIntegrationRecord: Record "CRM Integration Record" temporary)
    begin
        case RecordID.TableNo() of
            DATABASE::Customer:
                RemoveCoupledContactsForCustomer(RecordID, TempCRMIntegrationRecord);
        end;
        RemoveSingleCoupling(RecordID, TempCRMIntegrationRecord);
    end;

    local procedure RemoveSingleCoupling(RecordID: RecordID; var TempCRMIntegrationRecord: Record "CRM Integration Record" temporary)
    var
        CRMIntegrationRecord: Record "CRM Integration Record";
    begin
        CRMIntegrationRecord.RemoveCouplingToRecord(RecordID);

        TempCRMIntegrationRecord := CRMIntegrationRecord;
        TempCRMIntegrationRecord.Skipped := false;
        if TempCRMIntegrationRecord.Insert() then;
    end;

    local procedure RemoveCoupledContactsForCustomer(RecordID: RecordID; var TempCRMIntegrationRecord: Record "CRM Integration Record" temporary)
    var
        Contact: Record Contact;
        ContBusRel: Record "Contact Business Relation";
        Customer: Record Customer;
        CRMAccount: Record "CRM Account";
        CRMContact: Record "CRM Contact";
        CRMIntegrationRecord: Record "CRM Integration Record";
        TempContact: Record Contact temporary;
        CRMID: Guid;
    begin
        // Convert the RecordID into a Customer
        Customer.Get(RecordID);

        // Get the Company Contact for this Customer
        ContBusRel.SetCurrentKey("Link to Table", "No.");
        ContBusRel.SetRange("Link to Table", ContBusRel."Link to Table"::Customer);
        ContBusRel.SetRange("No.", Customer."No.");
        if ContBusRel.FindFirst() then begin
            // Get all Person Contacts under it
            Contact.SetCurrentKey("Company Name", "Company No.", Type, Name);
            Contact.SetRange("Company No.", ContBusRel."Contact No.");
            Contact.SetRange(Type, Contact.Type::Person);
            if Contact.FindSet() then begin
                // Count the number of Contacts coupled to CRM Contacts under the CRM Account the Customer is coupled to
                CRMIntegrationRecord.FindIDFromRecordID(RecordID, CRMID);
                if CRMAccount.Get(CRMID) then begin
                    repeat
                        if CRMIntegrationRecord.FindIDFromRecordID(Contact.RecordId(), CRMID) then begin
                            CRMContact.Get(CRMID);
                            if CRMContact.ParentCustomerId = CRMAccount.AccountId then begin
                                TempContact.Copy(Contact);
                                TempContact.Insert();
                            end;
                        end;
                    until Contact.Next() = 0;

                    // If any, query for breaking their couplings
                    if TempContact.Count() > 0 then
                        if Confirm(StrSubstNo(RemoveCoupledContactsUnderCustomerQst, TempContact.Count(), CRMProductName.FULL())) then begin
                            TempContact.FindSet();
                            repeat
                                RemoveSingleCoupling(TempContact.RecordId(), TempCRMIntegrationRecord);
                            until TempContact.Next() = 0;
                        end;
                end;
            end;
        end;
    end;
}

