module Pdf

open Spire.Xls
open System.IO

let convertToPdf pathDocument =

    use workbook = new Workbook()
    workbook.LoadFromFile(pathDocument)
    let sheet = workbook.Worksheets[0]

    let pdfFilePath = Path.ChangeExtension(pathDocument, "pdf")
    sheet.SaveToPdf(pdfFilePath)

    workbook.Dispose()
