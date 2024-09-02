using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

void RemplacerTexteDansExcel(string cheminDocument, Dictionary<string, string> valeurs)
{
    // Ouvrir le document Excel
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(cheminDocument, true))
    {
        // Obtenir la première feuille de calcul
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        // Parcourir toutes les lignes et cellules de la feuille
        foreach (Row row in sheetData.Elements<Row>().Where(row => row.RowIndex >= 1 && row.RowIndex <= 40))
        {
            //foreach (Cell cell in row.Elements<Cell>().Where(cell => cell.CellMetaIndex >= 1 && cell.CellMetaIndex <= 20))
            foreach (Cell cell in row.Elements<Cell>().Take(10))
            {
                string cellValue = GetCellValue(document, cell);

                foreach (var key in valeurs.Keys)
                {
                    // Remplacer les balises par les valeurs correspondantes
                    if (cellValue != null && cellValue.Contains(key))
                    {
                        cellValue = cellValue.Replace(key, valeurs[key]);
                        cell.CellValue = new CellValue(cellValue);
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    }
                }
            }
        }

        // Sauvegarder les modifications
        worksheetPart.Worksheet.Save();
    }
}

// Fonction utilitaire pour obtenir la valeur d'une cellule
string GetCellValue(SpreadsheetDocument document, Cell cell)
{
    string value = cell.CellValue?.InnerText;

    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    {
        return document.WorkbookPart.SharedStringTablePart.SharedStringTable
            .Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
    }

    return value;
}


var cheminDocument = "C:\\Users\\FamilleTE\\Downloads\\quittance-de-loyer-2024 - Copie - Copie.xlsx";
var valeurs = new Dictionary<string, string>
{
    { "{{Proprietaire.Prenom}}", "Jerome" },
    { "{{Proprietaire.Nom}}", "TE" },
    { "{{Proprietaire.Rue}}", "03 Rue Turgot" },
    { "{{Proprietaire.CodePostalVille}}", "75009 Paris" },
    { "{{Locataire.Prenom}}", "Julien" },
    { "{{Locataire.Nom}}", "Proust" },
    { "{{Locataire.Rue}}", "7 Rue Camille Flammarion " },
    { "{{Locataire.CodePostalVille}}", "91260 Juvisy-sur-Orge" },
    { "{{Loyer}}", "600" },
    { "{{Charges}}", "90" },
    { "{{Total}}", "690" },
};

RemplacerTexteDansExcel(cheminDocument, valeurs);
