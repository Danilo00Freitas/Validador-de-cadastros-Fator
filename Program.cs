using ConsolidadorFatorNelson.Automation;
using ConsolidadorFatorNelson.Menu;
using ConsolidadorFatorNelson.ExcelInteraction;

namespace ConsolidadorFatorNelson
{
    internal class Program

    {
    
        static void Main(string[] args)
        {
            var mainMenu = new MainMenu();
            mainMenu.DisplayTitle();
            var sheetPath = mainMenu.getSheetPath();    
            mainMenu.StartOperation(sheetPath);
            mainMenu.StopOperation();   
        }
    }
}
