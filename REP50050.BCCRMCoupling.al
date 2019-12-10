report 50500 "D365_BC CRM Coupling"
{
    UsageCategory = Administration;
    ApplicationArea = All;
    Caption = 'BC CRM Coupling';
    ProcessingOnly = true;

    dataset
    {
        dataitem(Integer; Integer)
        {
            DataItemTableView = SORTING(Number) WHERE(Number = CONST(1));
            trigger OnPreDataItem()
            begin
                //Reading Excel File
                ReadExcelBook();
            end;

            trigger OnAfterGetRecord()
            begin
                ExcelBuf.SetFilter("Row No.", '>%1', ExcelRowstoSkip);
                if ExcelBuf.FindSet() then
                    repeat
                        if RowNo <> ExcelBuf."Row No." then
                            if ExcelBuf."Row No." > (ExcelRowstoSkip + 1) then
                                ProcessData();
                        RowNo := ExcelBuf."Row No.";

                        case ExcelBuf."Column No." of
                            1:
                                NavPrimeKey := copystr(ExcelBuf."Cell Value as Text", 1, 20);
                            3:
                                begin
                                    CRMTempguidtext := copystr(ExcelBuf."Cell Value as Text", 1, 1024);
                                    Evaluate(CRMEntityGUID, CRMTempguidtext);
                                end;
                            5:
                                begin
                                    CreateNewCode := Copystr(Uppercase(ExcelBuf."Cell Value as Text"), 1, 5);
                                    if CreateNewCode IN ['TRUE', 'YES'] then
                                        CreateNew := true;
                                end;
                        end;
                    until ExcelBuf.Next() = 0;
                ProcessData();
            end;
        }
    }

    requestpage
    {
        layout
        {
            area(Content)
            {
                group(GroupName)
                {
                    field(FileName; ExcelFileName)
                    {
                        ApplicationArea = All;
                        Caption = 'Select File Name';
                        trigger OnAssistEdit()
                        begin
                            Clear(ExcelFileName);
                            if UploadIntoStream('Upload Excel File', 'C:\TEMP', 'All Files (*.*)|*.*', ExcelFileName, exInstream) THEN
                                ExcelBuf.Reset();
                        end;
                    }
                    field(SheetName; ExcelSheetName)
                    {
                        ApplicationArea = All;
                        Caption = 'Select Sheet Name';
                        trigger OnAssistEdit()
                        begin
                            if ExcelFileName = '' then
                                Error('First select Excel file');

                            Clear(ExcelSheetName);
                            ExcelSheetName := ExcelBuf.SelectSheetsNameStream(exInstream);

                        end;
                    }
                    field(NAVCRMCoupleEntity; NAVCRMCoupleEntity)
                    {
                        ApplicationArea = All;
                        Caption = 'Select Entity for Coupling Data';
                    }
                    field(RowstoSkip; ExcelRowstoSkip)
                    {
                        ApplicationArea = All;
                        Caption = 'Row to Skip';
                        MinValue = 1;
                    }
                }
            }
        }
    }

    var
        ExcelBuf: Record "Excel Buffer" temporary;
        Currency: record Currency;
        UOM: Record "Unit of Measure";
        Customer: Record Customer;
        Contact: Record Contact;
        Item: Record Item;
        Resource: Record Resource;
        CPG: Record "Customer Price Group";
        Opportunity: Record Opportunity;
        CRMSystemuser: Record "CRM Systemuser";
        CRMAccount: Record "CRM Account";
        CRMTransactioncurrency: Record "CRM Transactioncurrency";
        CRMContact: Record "CRM Contact";
        CRMUomSchedule: Record "CRM Uomschedule";
        CRMProduct: Record "CRM Product";
        CRMPricelevel: Record "CRM Pricelevel";
        CRMOpportunity: Record "CRM Opportunity";
        TempCRMSystemuser: Record "CRM Systemuser" temporary;
        ExcelRowstoSkip: Integer;
        RowNo: Integer;
        NavPrimeKey: Code[20];
        CRMEntityGUID: Guid;
        CRMTempguidtext: Text;
        CreateNewCode: Text[5];
        CreateNew: Boolean;
        NAVCRMCoupleEntity: Option " ","Salesperson/Purchaser",Currency,"Unit of Measure",Contact,Customer,Item,Resource,"Customer Price Group","Sales Price",Opportunity;
        ExcelFileName: Text;
        ExcelSheetName: Text;
        exInstream: InStream;

    trigger OnInitReport();
    begin
        ExcelRowstoSkip := 1;
    end;

    trigger OnPreReport();
    begin
        CODEUNIT.RUN(CODEUNIT::"CRM Integration Management");

        ExcelBuf.DeleteAll();
        if (ExcelFileName = '') or (ExcelSheetName = '') then
            Error('Either the Filename or the sheetname not selected');
        ExcelBuf.Reset();

        if NAVCRMCoupleEntity = NAVCRMCoupleEntity::" " then
            Error('Plase select Entity for Coupling Data');
    end;

    trigger OnPostReport();
    begin
        Message('Imported');
    end;

    local procedure ReadExcelBook()
    begin
        ExcelBuf.OpenBookStream(exInstream, ExcelSheetName);
        ExcelBuf.ReadSheet();
    end;

    local procedure ClearVar()
    begin
        Clear(NavPrimeKey);
        Clear(CRMTempguidtext);
        Clear(CRMEntityGUID);
        clear(CreateNewCode);
        Clear(CreateNew);
    end;

    local procedure InsertUpdateTempCRMSystemUser(SalespersonCode: Code[20]; SyncNeeded: Boolean)
    begin
        // FirstName is used to store coupled/ready to couple Salesperson
        // IsSyncWithDirectory is used to mark CRM User for coupling
        if TempCRMSystemuser.Get(CRMEntityGUID) then begin
            if not TempCRMSystemuser.IsDisabled or SyncNeeded then begin
                TempCRMSystemuser.FirstName := SalespersonCode;
                TempCRMSystemuser.IsSyncWithDirectory := SyncNeeded;
                TempCRMSystemuser.IsDisabled := SyncNeeded;
                TempCRMSystemuser.Modify();
            end
        end else begin
            TempCRMSystemuser.SystemUserId := CRMEntityGUID;
            TempCRMSystemuser.FirstName := SalespersonCode;
            TempCRMSystemuser.IsSyncWithDirectory := SyncNeeded;
            TempCRMSystemuser.IsDisabled := SyncNeeded;
            TempCRMSystemuser.Insert();
        end;
    end;

    local procedure CleanDuplicateSalespersonRecords(SalesPersonCode: Code[20]; CRMUserId: Guid)
    begin
        TempCRMSystemuser.Reset();
        TempCRMSystemuser.SetRange(FirstName, SalesPersonCode);
        TempCRMSystemuser.SetFilter(SystemUserId, '<>' + Format(CRMUserId));
        if TempCRMSystemuser.FindFirst() then begin
            TempCRMSystemuser.IsDisabled := true;
            TempCRMSystemuser.FirstName := '';
            TempCRMSystemuser.Modify();
        end;
    end;

    local procedure ProcessData()
    Var
        SalespersonPurchaser: Record "Salesperson/Purchaser";
        CRMIntegrationRecord: Record "CRM Integration Record";
        CRMAutoCoupleMngmt: Codeunit "D365_CRM AutoCouple Mngmt";
        CRMIntegrationManagement: Codeunit "CRM Integration Management";
        OldRecordId: RecordId;
        Synchronize: Boolean;
        Direction: Option;
    begin
        case NAVCRMCoupleEntity of

            //SalesPerson Coupling
            NAVCRMCoupleEntity::"Salesperson/Purchaser":
                Begin
                    CRMSystemuser.get(CRMEntityGUID);
                    if CRMSystemuser.IsIntegrationUser and CRMSystemuser.IsDisabled and not CRMSystemuser.IsLicensed then
                        CurrReport.Skip();

                    if not CreateNew then begin
                        if NavPrimeKey <> '' then begin
                            SalespersonPurchaser.Get(NavPrimeKey);
                            InsertUpdateTempCRMSystemUser(SalespersonPurchaser.Code, true);
                            CleanDuplicateSalespersonRecords(SalespersonPurchaser.Code, CRMEntityGUID);
                        end else
                            InsertUpdateTempCRMSystemUser('', true);

                        //Coupling & Synching
                        //if not CreateNew then begin
                        TempCRMSystemuser.Reset();
                        //TempCRMSystemuser.SetRange(IsSyncWithDirectory, true);
                        if TempCRMSystemuser.FindSet() then
                            repeat
                                if TempCRMSystemuser.FirstName <> '' then begin
                                    SalespersonPurchaser.Get(TempCRMSystemuser.FirstName);
                                    CRMIntegrationManagement.CoupleCRMEntity(
                                    SalespersonPurchaser.RecordId(), TempCRMSystemuser.SystemUserId, Synchronize, Direction);
                                end else begin
                                    CRMIntegrationRecord.FindRecordIDFromID(
                                    TempCRMSystemuser.SystemUserId, DATABASE::"Salesperson/Purchaser", OldRecordId);
                                    CRMAutoCoupleMngmt.RemoveCoupling(OldRecordId);
                                end;
                            until TempCRMSystemuser.Next() = 0;
                        TempCRMSystemuser.ModifyAll(IsSyncWithDirectory, false);
                        TempCRMSystemuser.ModifyAll(IsDisabled, false);
                        TempCRMSystemuser.DeleteAll();
                    end else
                        CRMIntegrationManagement.CreateNewRecordsFromCRM(CRMSystemuser);
                End;
            // Currency Coupling
            NAVCRMCoupleEntity::Currency:
                Begin
                    Currency.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMTransactioncurrency.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Currency.RecordId(), CRMTransactioncurrency.CurrencyName, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Currency.RecordId(), Currency.Description, CreateNew);
                End;

            // Unit of Measure Coupling
            NAVCRMCoupleEntity::"Unit of Measure":
                Begin
                    UOM.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMUomSchedule.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(UOM.RecordId(), CRMUomSchedule.Name, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(UOM.RecordId(), UOM.Description, CreateNew);
                End;

            // Customer Coupling
            NAVCRMCoupleEntity::Customer:
                Begin
                    Customer.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMAccount.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Customer.RecordId(), CRMAccount.Name, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Customer.RecordId(), customer.Name, CreateNew);
                End;

            // Contact Coupling
            NAVCRMCoupleEntity::Contact:
                Begin
                    Contact.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMContact.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Customer.RecordId(), CRMContact.FirstName, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Customer.RecordId(), contact.Name, CreateNew);
                End;

            // Item Coupling
            NAVCRMCoupleEntity::Item:
                Begin
                    Item.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMProduct.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Item.RecordId(), CRMProduct.Name, CreateNew)
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Item.RecordId(), Item.Description, CreateNew);
                End;

            // Resource Coupling
            NAVCRMCoupleEntity::Resource:
                Begin
                    Resource.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMProduct.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Resource.RecordId(), CRMProduct.Name, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Resource.RecordId(), Resource.Name, CreateNew);
                End;

            // Customer Price Group Coupling
            NAVCRMCoupleEntity::"Customer Price Group":
                Begin
                    CPG.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMPricelevel.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(CPG.RecordId(), CRMPricelevel.Name, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(CPG.RecordId(), CPG.Description, CreateNew);
                End;

            // Opportunity Coupling
            NAVCRMCoupleEntity::Opportunity:
                Begin
                    Opportunity.get(NavPrimeKey);
                    if not CreateNew then BEGIN
                        CRMOpportunity.get(CRMEntityGUID);
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Opportunity.RecordId(), CRMOpportunity.Name, CreateNew);
                    end else
                        CRMAutoCoupleMngmt.DefineAutoCoupling(Opportunity.RecordId(), Opportunity.Description, CreateNew);
                End;
        end;

        ClearVar();
    end;
}