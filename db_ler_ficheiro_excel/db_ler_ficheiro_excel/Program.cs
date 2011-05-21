using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;

namespace db_ler_ficheiro_excel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //isto tem que ser melhorado aqui. mas deixamos essa parte para quando incluirmos a parte gráfica. 
                // o user deve escolher o ficheiro .xls/xlsx que quiser
                DirectoryInfo directorio = new DirectoryInfo(@"C:\Users\Alfredo Pinheiro\Desktop\db_linha_de_comandos\db_linha_de_comandos\bin\Debug\blah.xlsx");
                string output = @"C:\Users\Alfredo Pinheiro\Desktop\db_linha_de_comandos\db_linha_de_comandos\bin\Debug\blah.xlsx";

                Console.WriteLine("A extrair dados...");
                ler_ficheiro_excel.ler_ficheiro(output);
            }

            catch(Exception ex)
            {
                Console.WriteLine("Erro {0}", ex.Message);
            }

            Console.WriteLine();
            Console.Read();
        }
    }
}
