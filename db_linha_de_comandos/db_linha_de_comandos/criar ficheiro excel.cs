using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Xml;
using OfficeOpenXml.Style;  //necessário para editar as cores, etc.
using System.Drawing;

namespace db_linha_de_comandos
{
    class criar_ficheiro_excel
    {
        public static string criar_ficheiro(DirectoryInfo ficheiro_excel)
        {
           FileInfo _ficheiro_novo = new FileInfo(ficheiro_excel.FullName + @"\blah.xlsx");
           if (_ficheiro_novo.Exists)
           {
               _ficheiro_novo.Delete(); //só para ter a certeza que estamos a trabalhar com o ficheiro certo
               _ficheiro_novo = new FileInfo(ficheiro_excel.FullName + @"\blah.xlsx");
           }

            //aqui começa a verdadeira edição do ficheiro xlsx
           using (ExcelPackage _package_ = new ExcelPackage(_ficheiro_novo))
           {
               //uma coisa básica só para exemplificar, pode ser util no futuro o básico
               //criar uma novo folha dentro de um "workbook"
               ExcelWorksheet ws = _package_.Workbook.Worksheets.Add("Alunos");     //nome da nossa folha do excel

               //configurar os campos
               ws.Cells[1, 1].Value = "Nome";   //estou a escrever coisas random :P
               ws.Cells[1, 2].Value = "Ano";
               ws.Cells[1, 3].Value = "Turma";
               ws.Cells[1, 4].Value = "Idade";
               ws.Cells[1, 5].Value = "Nota final";     //LOL
               ws.Cells[1, 6].Value = "Fórmula random";


               //dentro desses campos adicionar agora os detalhes de cada um
               ws.Cells["A2"].Value = "André Pinheiro";
               ws.Cells["B2"].Value = "11º";
               ws.Cells["C2"].Value = "CT1";
               ws.Cells["D2"].Value = 16;
               ws.Cells["E2"].Value = 21;

               //outro aluno
               ws.Cells["A3"].Value = "Rafael Almeida";
               ws.Cells["B3"].Value = "11º";
               ws.Cells["C3"].Value = "CT2";
               ws.Cells["D3"].Value = 16;
               ws.Cells["E3"].Value = 10;   //:P

               //sei lá... agora.. por exemplo, adicionar uma fórmula qualquer 
               //tem que ser a frente do E claro..e entre as células F2:F3
               ws.Cells["F2:F3"].Formula = "(D2*0.20)+(E2*0.80)";   //adiciona automaticamente aos próximos..

               //está tudo feito. vamos passar para a próxima fase..
               //isto não é necessário, é só para ficar mais bonito(ou feio)(cópia)
               using (var range = ws.Cells[1, 1, 1, 6])     //ws.cells[row, col, toRow, toCol(no nosso caso F)]
               {
                   range.Style.Font.Bold = true;    //não interessa esta parte do código
                   range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                   range.Style.Fill.BackgroundColor.SetColor(Color.Red);
                   range.Style.Font.Color.SetColor(Color.White);
               }

               ws.Cells["A4:F4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;    //design..
               ws.Cells["D2:D5"].Style.Numberformat.Format = "###";
               ws.Cells["E2:E3"].Style.Numberformat.Format = "###";


               ws.Cells["A1:F6"].AutoFilter = true;     //ainda não percebi o que isto faz :S

               ws.Column(1).Width = 15;
               ws.Column(2).Width = 7;
               ws.Column(5).Width = 15;
               ws.Column(6).Width = 15;

               //mais design(cópia). Dá para se perceber muito bem isto.
               ws.HeaderFooter.oddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Alunos";
               ws.HeaderFooter.oddFooter.RightAlignedText = string.Format("Página {0} de {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
               ws.HeaderFooter.oddFooter.CenteredText = ExcelHeaderFooter.SheetName;

               //para mudar o modo de visualização. não é do nosso interesse para já por isso vou deixar comentado
               ws.View.PageLayoutView = true;    

               //pronto, só falta isto que fica sempre bem se quisermos :P
               _package_.Workbook.Properties.Author = ".....";
               _package_.Workbook.Properties.Title = "criar_ficheiro_excel";
               _package_.Workbook.Properties.Comments = "(Nada para dizer)";


               //E agora claro, temos que gravar no directório actual
               _package_.Save();    //podemos mudar isto facilmente. serve só de base para o resto
           }

           return _ficheiro_novo.FullName;
        }
    }
}
