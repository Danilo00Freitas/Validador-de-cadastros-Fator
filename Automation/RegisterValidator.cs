using DocumentFormat.OpenXml.Bibliography;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsolidadorFatorNelson.Automation
{
    internal class RegisterValidator
    {
        private string _baseDirectory;
        private string _chromePath;
        private string _chromeDriverPath;
        private ChromeOptions _options;
        private IWebDriver _driver;

        public RegisterValidator() 
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--headless");
            options.BinaryLocation = _chromePath;
            _options = options;
            var chromeDriver = new ChromeDriver(_chromeDriverPath, _options);
            _driver = chromeDriver;
        }

        public IReadOnlyDictionary<string, String> validateCnpj(IReadOnlyCollection<String> cnpjList)
        {
            if (_driver == null)
            {
                throw new InvalidOperationException("The WebDriver instance is null.");
            }

            //Configurando esperas
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
            Dictionary<string, String> cnpjRegisterStatusDict = new Dictionary<string, String>();
            int totalClientes = cnpjList.Count;
            int clientesVerificados = 0;


            foreach (String cnpj in cnpjList)
            {
                try 
                {
                    // Navegando até o site
                    _driver.Navigate().GoToUrl("https://canaldigital.fatorconnect.com.br/cadastro/validar-dados");
                } 
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao acessar a página principal \n{ex.Message}");
                    continue;
                }

                try 
                {
                    //Procurando e preenchendo o campo "Número de CNPJ"
                    var numeroDoCnpjField = wait.Until(d => d.FindElement(By.XPath("//*[@id=\"mat-input-0\"]")));
                    numeroDoCnpjField.SendKeys(cnpj);
                    var bodyElement = _driver.FindElement(By.TagName("body"));
                    bodyElement.Click();
                } 
                catch (Exception ex) 
                {
                    Console.WriteLine($"Erro ao procurar e preencher o campo NÚMERO DO CNPJ \n{ex.Message}");
                    continue;
                }
                
                try 
                {
                    var cnpjJaCadastradoField = wait.Until(d => d.FindElement(By.XPath("//*[@id=\"mat-dialog-0\"]")));
                    String registered = "Registrado";


                    if (!cnpjRegisterStatusDict.ContainsKey(cnpj))
                    {
                        cnpjRegisterStatusDict.Add(cnpj, registered);
                    }
                }
                catch
                {
                    String registered = "Não registrado";

                    if (!cnpjRegisterStatusDict.ContainsKey(cnpj))
                    {
                        cnpjRegisterStatusDict.Add(cnpj, registered);
                    }
                }

                clientesVerificados++;

                // Calcula quantos clientes faltam para verificar
                int clientesFaltando = totalClientes - clientesVerificados;

                // Exibe o progresso atual em uma única linha
                Console.Write($"\rClientes verificados: {clientesVerificados} / {totalClientes}. Faltam {clientesFaltando} clientes.");
            
        }
            Console.WriteLine();
            return cnpjRegisterStatusDict.AsReadOnly();
        }
    }
}
