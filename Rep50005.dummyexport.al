report 50102 CustomerreportExcel
{
    ProcessingOnly = true;
    UseRequestPage = true;
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All, Basic, Suite;
    Caption = 'Customer  Export';
    dataset
    {

        dataitem(Customer; Customer)
        {
            
            RequestFilterFields = "No.";
            trigger OnPreDataItem();
            var
            //Excel : Integer;
            begin
                ExcelBuffer_gRec.DeleteAll();
                ExcelBuffer_gRec.NewRow();
                ExcelBuffer_gRec.AddColumn('Customer Balance Sheet', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.NewRow();
                ExcelBuffer_gRec.AddColumn('Customer Code', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn('Customer Name', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn('Customer Balance', false, '', true, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                //ExcelBuffer_gRec.NewRow;
            end;

            trigger OnAfterGetRecord();
            begin
                Customer.CalcFields("Balance (LCY)");
                ExceLBuffer_gRec.NewRow();
                ExcelBuffer_gRec.AddColumn("No.", false, '', False, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn(Name, false, '', False, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);
                ExcelBuffer_gRec.AddColumn("Balance (LCY)", false, '', false, false, false, '', ExcelBuffer_gRec."Cell Type"::Text);

            end;

            trigger OnPostDataItem();
            begin


            end;
        }


    }



    var
        ExcelBuffer_gRec: Record "Excel Buffer" temporary;
        //GeneralLedgersetupExt_gRec: Record "General Ledger Setup Extension";

    trigger OnInitReport();
    begin

    end;

    trigger OnPreReport();
    begin
        ExcelBuffer_gRec.DeleteAll;
    end;

    trigger OnPostReport();
    begin
        //GeneralLedgersetupExt_gRec.get;
        ExcelBuffer_gRec.CreateNewBook('Customer');
        ExcelBuffer_gRec.WriteSheet('Customer', CompanyName, UserId);
        ExcelBuffer_gRec.CloseBook;
        ExcelBuffer_gRec.OpenExcel;
    end;
}

