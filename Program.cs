using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace AcePlanilha
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var info = ReadXls();

            foreach (var item in info)
            {
                Console.WriteLine($"Nome: {item.Nome}\nPontuacao: {item.Pontuacao}");
            }

            Console.WriteLine();

            int totalPontuacao = info.Count();
            Console.WriteLine($"Quantidade Total: {totalPontuacao}");

            int pontuacoesMenoresQueSeis = info.Count(i => i.Pontuacao <= 6.25);
            Console.WriteLine($"Quantidade de pontuações menores que 6.25: {pontuacoesMenoresQueSeis}");

            int maiorQueseis = info.Count(i => i.Pontuacao > 6.25);
            Console.WriteLine($"Quantidade de pontuações maiores que 6.25: {maiorQueseis}");

            int igualSeis = info.Count(i => i.Pontuacao == 6.25);
            Console.WriteLine($"Quantidade de pontuações igual a 6.25: {igualSeis}");

            int maiusculas = info.Count(i => i.Nome == i.Nome.ToUpper()); // Conta nomes em caixa alta

            Console.WriteLine($"Quantidade de nomes em caixa alta: {maiusculas}");

            Console.ReadKey();
        }

        private static List<Info> ReadXls()
        {
            var response = new List<Info>();

            FileInfo existingFile = new FileInfo("C:\\Users\\Bruno\\OneDrive\\Área de Trabalho\\AprovadosACE.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Acrescentar2"];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 2; row <= rowCount; row++)
                {
                    var info = new Info();
                   
                    if (worksheet.Cells[row, 1].Value != null)
                    {
                        info.Nome = worksheet.Cells[row, 1].Value.ToString();
                    }

                    if (worksheet.Cells[row, 2].Value != null)
                    {
                        info.Pontuacao = Convert.ToDouble(worksheet.Cells[row, 4].Value);
                    }

                    response.Add(info);
                }

                return response;
            }
        }
    }

    public class Info
    {
        public string Nome { get; set; }
        public double Pontuacao { get; set; }
    }
}
