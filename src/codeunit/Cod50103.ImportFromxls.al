codeunit 50103 "Import From xls"
{

    procedure SetOrderNo(p_OrderNo: Code[20])
    begin
        OrderNo := p_OrderNo;
    end;

    procedure ReadXls()
    var
        Tempblob: Record TempBlob temporary;
        FileName: Text;
        Instr: InStream;
    begin
        ExcelBuffer.Reset();
        ExcelBuffer.DeleteAll();

        Tempblob.Blob.CreateInStream(Instr);

        UploadIntoStream('', '', '', FileName, Instr);
        ExcelBuffer.OpenBookStream(Instr, 'Sheet1');
        ExcelBuffer.ReadSheet();
        ParseXls();
    end;

    local procedure ParseXls()
    var
        Item: Record Item;
        FirstRowNo: Integer;
        LastRowNo: Integer;
        i: Integer;
    begin
        FirstRowNo := 0;
        ExcelBuffer.SetRange("Column No.", 2);
        ExcelBuffer.SetFilter("Cell Value as Text", '<>%1', '');
        if ExcelBuffer.FindFirst() then
            repeat
                If Item.Get(ExcelBuffer."Cell Value as Text") then
                    FirstRowNo := ExcelBuffer."Row No.";
            until (ExcelBuffer.Next() = 0) or (FirstRowNo > 0);

        ExcelBuffer.Reset();
        ExcelBuffer.SetFilter("Cell Value as Text", '<>%1', '');
        if ExcelBuffer.FindLast() then
            LastRowNo := ExcelBuffer."Row No.";

        ExcelBuffer.Reset();
        LineNo := GetSalesLineNo();

        For i := FirstRowNo to LastRowNo do
            InsertSalesLine(i);

    end;

    local procedure InsertSalesLine(RowNo: Integer)
    var
        SalesLine: Record "Sales Line";

        ItemNo: code[20];
        Qty: Decimal;
    begin
        If ExcelBuffer.Get(RowNo, 2) then
            ItemNo := ExcelBuffer."Cell Value as Text";

        If ExcelBuffer.Get(RowNo, 4) then
            Evaluate(Qty, ExcelBuffer."Cell Value as Text");

        If ItemNo <> '' then begin
            LineNo += 10000;
            SalesLine.Init();
            SalesLine.Validate("Document Type", SalesLine."Document Type"::Order);
            SalesLine.Validate("Document No.", OrderNo);
            SalesLine."Line No." := LineNo;
            SalesLine.Validate(Type, SalesLine.Type::Item);
            SalesLine.Validate("No.", ItemNo);
            SalesLine.Validate(Quantity, Qty);
            SalesLine.Insert();

        end;
    end;

    local procedure GetSalesLineNo(): Integer
    var
        SalesLine: Record "Sales Line";
    begin
        SalesLine.Setrange("Document Type", SalesLine."Document Type"::Order);
        SalesLine.SetRange("Document No.", OrderNo);
        If SalesLine.FindLast() then
            exit(SalesLine."Line No.");

        exit(0);
    end;

    var
        ExcelBuffer: Record "Excel Buffer" temporary;
        OrderNo: Code[20];
        LineNo: Integer;
}