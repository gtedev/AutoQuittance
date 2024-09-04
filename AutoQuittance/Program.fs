let config = Config.getConfig ()

config.templates
|> List.iter (fun t ->
    Excel.replaceTextInExcel t
    Pdf.convertToPdf t.outputFile)
