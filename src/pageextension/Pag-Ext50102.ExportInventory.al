pageextension 50102 "ExportInventory" extends "Location List" //MyTargetPageId
{

    actions
    {
        addlast(Processing)
        {
            action(ExportInventory)
            {
                ApplicationArea = All;
                RunObject = codeunit Loc_to_xls;
                Image = Export;
            }
        }
    }
}