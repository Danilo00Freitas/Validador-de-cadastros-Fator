using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorFatorNelson.SheetManager;

internal class SheetCreator
{
    private String _sheetPath;
    public SheetCreator(string sheetPath)
    {
        _sheetPath = sheetPath;
    }

    public void updateSheet(IReadOnlyDictionary<String,String> dictonary)
    {
        int rowAccounter = 2;
        try 
        {
            using (var workbook = new XLWorkbook(_sheetPath))
            {
                var worksheet = workbook.Worksheet(1);
                int cnpjColumnNumber = 5;
                int registerStatusColumnNumber = 11;

                foreach (var cell in worksheet.Column(cnpjColumnNumber).CellsUsed())
                {
                    String cellValue = cell.GetString().Trim().Replace(" ","").ToLower();
                    if (!string.IsNullOrEmpty(cellValue) && !cellValue.Equals("cnpj".ToLower().Trim().Replace(" ",""), StringComparison.OrdinalIgnoreCase))
                    {
                        if (dictonary.ContainsKey(cellValue))
                        {
                            String status = dictonary[cellValue];
                            worksheet.Cell(rowAccounter, registerStatusColumnNumber).Value = status;
                            rowAccounter++;
                        }
                        else
                        {
                            rowAccounter++;
                        }

                    }
                    else
                    {
                        rowAccounter++;
                    }
                }
                workbook.Save();
            }
        } 
        catch (Exception ex) 
        { 
            Console.WriteLine($"Erro ao criar atualizar a planilha{ex.Message}");
        }
    }
}
