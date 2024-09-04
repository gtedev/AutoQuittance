module Excel

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open System.IO
open Config

let (|HasCellValue|_|) (cell: Cell) =
    if (cell.CellValue = null) then None else Some()

let (|IsSharedString|_|) (cell: Cell) =
    if (cell.DataType.Value <> CellValues.SharedString) then
        None
    else
        Some(int cell.CellValue.InnerText)

let tryGetCellValue (doc: SpreadsheetDocument) (cell: Cell) =
    match cell with
    | HasCellValue & IsSharedString cellIndex ->

        // Get cellValue from the SharedString table
        // with the help of the cellIndex
        doc.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
        |> Seq.tryItem cellIndex
        |> Option.map (fun sharedStringItem -> cell, sharedStringItem.InnerText)

    | _ -> None

let replaceText (cellValue: string) (key: string) (value: string) =
    if cellValue.Contains(key) then
        cellValue.Replace(key, value)
    else
        cellValue

let updateCellValue (cell: Cell) (newCellValue: string) =
    cell.CellValue <- new CellValue(newCellValue)
    cell.DataType <- new EnumValue<CellValues>(CellValues.String)

let replaceTextInExcel (t: Template) =
    // 1. Copy the template file
    File.Copy(t.pathFile, t.outputFile)

    // 2. Then, open it
    use document = SpreadsheetDocument.Open(t.outputFile, true)

    let workbookPart = document.WorkbookPart
    let sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>()
    let worksheetPart = workbookPart.GetPartById(sheet.Id) :?> WorksheetPart
    let sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()

    // 3. Then, replace values in the newly copied file
    sheetData.Elements<Row>()
    |> Seq.collect (fun row -> row.Elements<Cell>())
    |> Seq.choose (tryGetCellValue document)
    |> Seq.iter (fun (cell, cellValue) ->
        let newCellValue =
            t.properties
            |> Map.toList
            |> List.map (fun (key, value) -> "{{" + key + "}}", value)
            |> Map<string, string>
            |> Map.fold replaceText cellValue

        updateCellValue cell newCellValue
        ())

    worksheetPart.Worksheet.Save()
    document.Save()

    // To make sure, the changes are flushed to the file.
    document.Close()
