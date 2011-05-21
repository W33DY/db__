using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;

namespace db_ler_ficheiro_excel
{
    class ler_ficheiro_excel
    {
        public static void ler_ficheiro(string pasta_ficheiro)
        {
            FileInfo _ficheiro = new FileInfo(pasta_ficheiro);

            using (ExcelPackage _package = new ExcelPackage(_ficheiro))  //acho que já entendeste o que isto faz
            {
                //dentro do workbook vamos apenas utilizar a primeira(e unica) folha que fizemos anteriormente
                //a folha com o nome "Alunos"
                ExcelWorksheet ws = _package.Workbook.Worksheets[1];

                int coluna = 2;     //a coluna do ano de escolaridade
                for (int linha = 2; linha < 4; linha++)
                    Console.WriteLine("Célula({0}, {1}). Valor = {2}", linha, coluna, ws.Cells[linha, coluna].Value);
            }

            Console.WriteLine("Feito...Dados extraidos");
        }
    }
}
