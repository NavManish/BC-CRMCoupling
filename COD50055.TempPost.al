codeunit 50055 TempPost
{
    
    trigger OnRun()
    begin
        
    end;
    
    procedure PostPaymentJnl()
    var
        GenJnl : Record "Gen. Journal Line";
        PostBatch : Codeunit "Gen. Jnl.-Post Batch";
        Posted : Boolean;
    begin
        GenJnl.Reset();
        GenJnl.SetRange("Journal Template Name",'CASHRCPT');
        GenJnl.SetRange("Journal Batch Name",'GENERAL');
        if not GenJnl.IsEmpty() then
          Posted := PostBatch.Run(GenJnl);
    end;
    var
        myInt: Integer;
}