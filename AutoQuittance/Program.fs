let config = Config.getConfig ()

config.templates |> List.iter Excel.replaceTextInExcel
