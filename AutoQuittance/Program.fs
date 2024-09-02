open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet


let tryGetCellValue  (doc: SpreadsheetDocument)(cell: Cell) = 
      let someCellValue = 
         if (cell.CellValue = null ) then None else Some (cell.CellValue.InnerText)

      if cell.DataType <> null && cell.DataType.Value = CellValues.SharedString then
          someCellValue
          |> Option.map (fun cellValue ->
                  doc.WorkbookPart
                    .SharedStringTablePart
                    .SharedStringTable
                    .Elements<SharedStringItem>() 
                    |> Seq.tryItem (int cellValue)
                    |> Option.map(fun item -> cell, item.InnerText ))
          |> Option.flatten

      else 
          None

let replaceText (cellValue: string) (key:string)  (value:string)  = 
        if cellValue.Contains(key) then  
                cellValue.Replace(key, value)
        else 
            cellValue

let replaceTextInExcel (pathFile: string) (values: Map<string, string> ) =
        let document = SpreadsheetDocument.Open(pathFile, true)
        let workbookPart = document.WorkbookPart;
        let sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>()
        let worksheetPart = workbookPart.GetPartById(sheet.Id) :?> WorksheetPart
        let sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()

        sheetData.Elements<Row>()
        |> Seq.collect (fun row -> row.Elements<Cell>())
        |> Seq.choose (tryGetCellValue document)
        |> Seq.iter(fun (cell, cellValue) -> 
                        let newCellValue = values |> Map.fold(fun acc key value -> replaceText acc key value) cellValue
                        cell.CellValue <- new CellValue(newCellValue) 
                        cell.DataType <- new EnumValue<CellValues>(CellValues.String)
                        ()
          )

        worksheetPart.Worksheet.Save()
        document.Save()

        // To make sure, the changes are flushed to the real file.
        document.Close()

let pathDocument = "C:\\Users\\FamilleTE\\Downloads\\quittance-de-loyer-2024 - Copie - Copie.xlsx"

let values = 
    [|
    ( "{{Proprietaire.Prenom}}", "Gerard" )
    ( "{{Proprietaire.Nom}}", "TE" )
    ( "{{Proprietaire.Rue}}", "09 rue des pepinieres" )
    ( "{{Proprietaire.CodePostalVille}}", "75009 Paris" )
    ( "{{Locataire.Prenom}}", "Julien" )
    ( "{{Locataire.Nom}}", "Proust" )
    ( "{{Locataire.Rue}}", "7 Rue Camille Flammarion " )
    ( "{{Locataire.CodePostalVille}}", "91260 Juvisy-sur-Orge" )
    ( "{{Loyer}}", "600" )
    ( "{{Charges}}", "90" )
    ( "{{Total}}", "690" )|]  
      |> Map.ofArray 


replaceTextInExcel  pathDocument values