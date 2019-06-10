codeunit 50102 "Loc_to_xls"
{
    trigger OnRun()
    begin
        Export();
    end;

    local procedure Export()
    var
        Location: Record Location;
        ExportInventory: Report ExportInvToxls;
    begin
        if Location.FindFirst() then
            repeat
                Clear(ExportInventory);
                ExportInventory.SetLocation(Location.Code);
                //ExportInventory.SaveAsExcel('C:\Temp\' + Location.Code + '.xlsx');
                ExportInventory.SaveAsPdf('C:\Temp\' + Location.Code + '.pdf');
            until Location.Next() = 0;
    end;
}