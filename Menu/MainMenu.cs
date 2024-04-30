using ConsolidadorFatorNelson.Automation;
using ConsolidadorFatorNelson.ExcelInteraction;
using ConsolidadorFatorNelson.SheetManager;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;

namespace ConsolidadorFatorNelson.Menu
{
    internal class MainMenu
    {
        private CnpjListCreator _cnpjListCreator; 
        private RegisterValidator _registerValidator;
        private SheetCreator _sheetCreator;
        private IWebDriver _driver;

        public MainMenu()
        {
            _registerValidator = new RegisterValidator();   

        }

        public void DisplayTitle()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine("***************************************");
            Console.WriteLine("* Bem-vindo ao Consolidador de clientes da Fator! *");
            Console.WriteLine("***************************************");
            Console.ResetColor();
            Console.WriteLine();
        }

        public String getSheetPath()
        {
            Console.WriteLine("Informe o caminho da planilha de clientes");
            var path = Console.ReadLine()!;
            return path;
        }

        public void StartOperation(String path) 
        {
            _cnpjListCreator = new CnpjListCreator(path);
            _sheetCreator = new SheetCreator(path);

            //Criando lista com CNPJs para validação
            Console.WriteLine("CRIANDO A LISTA DE CNPJS");
            var cnpjList = _cnpjListCreator.getAllCnpj();
            Console.WriteLine($"Foram encontrados {cnpjList.Count()} para validação");

            //Validando os CNPJS no site e criando uma coleção de CNPJ x Status
            var cnpjListValidated = _registerValidator.validateCnpj(cnpjList);

            //Populando a planilha com os resultados
            _sheetCreator.updateSheet(cnpjListValidated);

            Console.WriteLine("FINALIZADO");

        }

        public void StopOperation()
        {
            // Verifica se o WebDriver foi inicializado
            if (_driver != null)
            {
                // Fecha o navegador e encerra o WebDriver
                _driver.Quit();
                _driver.Dispose();
            }

            Console.WriteLine("Aperte qualquer tecla para encerrar");
            Console.ReadKey();
            Environment.Exit(0);

        }
    }


}

