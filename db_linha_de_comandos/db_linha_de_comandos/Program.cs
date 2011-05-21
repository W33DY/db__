using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;  //não é necessário..
using System.IO;

namespace db_linha_de_comandos
{
    class Program
    {
        static void Main(string[] args)
        {
            //vamos usar um bloco try/catch para evitar erros
            try
            {
                string directorio = @"C:";
                DirectoryInfo _directorio = new DirectoryInfo(directorio);

                //usei um que existe sempre no pc, se não existir tem que dar erro obviamente
                if (!_directorio.Exists) throw new Exception("A pasta que seleccionou não existe ;)");

                Console.WriteLine("A criar folha do excel");
                string _folha_final = criar_ficheiro_excel.criar_ficheiro(_directorio);
                Console.WriteLine("Folha do excel criada com sucesso! na pasta {0}", _directorio);
                Console.WriteLine();
            }

            catch (Exception ex)
            {
                Console.WriteLine("Erro {0}", ex.Message);
            }
            Console.Read();
        }
    }
}
