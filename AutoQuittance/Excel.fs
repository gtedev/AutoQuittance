module Excel

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet
open System.IO
open System
open Config

let (|HasCellValue|_|) (cell: Cell) =
    if (cell.CellValue = null) then None else Some()

let (|IsSharedString|_|) (cell: Cell) =
    if (cell.DataType = null || cell.DataType.Value <> CellValues.SharedString) then
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
    File.Copy(t.pathFile, t.outputFile, true)

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

        let computedProperties = 
              t.displayComputedProperties
              |> List.choose(fun x ->
                  let today = DateTime.Now

                  match x with
                  | "StartDate" -> 
                        let starDate = DateTime(today.Year, today.Month, t.period.startDayInMonth)
                        Some ("StartDate", starDate.ToString("dd/MM/yyyy"))
                  | "EndDate" -> 
                        let starDate = DateTime(today.Year, today.Month, t.period.startDayInMonth)
                        let endDate = if t.period.startDayInMonth = 1 then
                                        DateTime(today.Year, today.Month, DateTime.DaysInMonth(today.Year, today.Month)) 
                                      else 
                                        starDate.AddMonths(1)
                        Some ("EndDate", endDate.ToString("dd/MM/yyyy"))

                  | "SentDate" ->
                        let senDate = DateTime(today.Year, today.Month, t.period.sendDayInMonth)
                        Some ("SentDate", senDate.ToString("dd/MM/yyyy")) 
                  | _ -> None
              )

        let newCellValue =
            t.displayCustomProperties
            |> Map.toList
            |> List.append computedProperties
            |> List.map (fun (key, value) -> "{{" + key + "}}", value)
            |> Map<string, string>
            |> Map.fold replaceText cellValue

        // Update only if there is new value to replace
        if(newCellValue <> cellValue) then
            updateCellValue cell newCellValue
        ())

    worksheetPart.Worksheet.Save()

    // To make sure, the changes are flushed to the file.
    document.Close()
