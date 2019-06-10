report 50102 "ExportInvToxls"
{
    UsageCategory = Administration;
    ApplicationArea = All;
    Caption = 'Export Inventory to xls';
    RDLCLayout = 'exporttoxls.rdl';

    dataset
    {
        dataitem(Item; Item)
        {
            CalcFields = Inventory;

            column(No_; "No.")
            {

            }
            column(Description; Description)
            {

            }
            column(Inventory; Inventory)
            {

            }

            trigger OnPreDataItem()
            begin
                Item.SetFilter("Location Filter", LocationCode);
                Item.SetFilter(Inventory, '>0');
                If Item.FindSet() then;
            end;
        }
    }

    procedure SetLocation(p_LocationCode: Code[10])
    begin
        LocationCode := p_LocationCode;
    end;

    var
        LocationCode: Code[10];
}