using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorFatorNelson.ExcelInteraction;

internal class CnpjListCreator
{

    private String _sheetPath;
    private List<string> _cnpjList = new List<string>();

    public CnpjListCreator(string sheetPath)
    {
        _sheetPath = sheetPath;
        Console.WriteLine($"PESQUISANDO PLANILHA NO CAMINHO: \n {_sheetPath}");
    }

    public IReadOnlyList<string> getAllCnpj()
    {
        try
        {
            using (var workbook = new XLWorkbook(_sheetPath))
            {
                var worksheet = workbook.Worksheet(1);
                int cnpjColumnNumber = 5;

                foreach (var cell in worksheet.Column(cnpjColumnNumber).CellsUsed())
                {
                    String rawCellValue = cell.GetString();
                    String cellValue = rawCellValue.ToLower().Trim().Replace(" ", "");
                    if (!string.IsNullOrEmpty(cellValue) && !cellValue.Equals("CNPJ", StringComparison.OrdinalIgnoreCase))
                    {
                        _cnpjList.Add(cellValue);
                    }
                    
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro no processo de abertura da planilha e captura dos CNPJS: \n{ex.Message}");

            return _cnpjList.AsReadOnly();

        }
        return _cnpjList.AsReadOnly();
    }
}