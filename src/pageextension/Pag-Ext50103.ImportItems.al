pageextension 50103 "Import Items" extends "Sales Order" //MyTargetPageId
{

    actions
    {
        addlast(Processing)
        {
            action(ImportFromXls)
            {
                ApplicationArea = All;

                Image = ImportExcel;

                trigger OnAction()
                var
                    ImportXls: codeunit "Import From xls";
                begin
                    ImportXls.SetOrderNo("No.");
                    ImportXls.ReadXls();
                end;
            }
        }
    }
}