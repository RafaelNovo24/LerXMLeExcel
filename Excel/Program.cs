using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Excel
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string filepath, returnMethod="";
            string resposta = Console.ReadLine();

            Console.WriteLine("**************** Bem vindo ao menu de testes Excel ****************");
            Console.WriteLine("Escolha a opção:");
            Console.WriteLine($@"1.-Importar Excel.");
            Console.WriteLine($"2.- Outros");
            Console.WriteLine($@"3.-Sair");
            resposta = Console.ReadLine();


            switch (resposta)
            {
                case "1":
                    filepath = ImportarExcel();
                   returnMethod= LerExcel(filepath);
                    Console.WriteLine(returnMethod);
                    Console.ReadLine();
                    break;

                case "2":

                    break;

                case "3":
                    break;

                default:
                    break;
            }
        }

        static string ImportarExcel()
        {
            string filePath = "";
            try
            {
                System.Windows.Forms.OpenFileDialog dia = new OpenFileDialog();

                //dia.InitialDirectory = $@"C:\";
                dia.Filter = "Excel Files| *.xls; *.xlsx; *.xlsm";
                //dia.RestoreDirectory = true;

                if (dia.ShowDialog() == DialogResult.OK)
                {
                    filePath = dia.FileName;
                    var filestream = dia.OpenFile();
                }
                return filePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro" + ex.Message);
                return filePath;
            }
        }

        private static string LerExcel(string filepath)
        {
            int linhaFinal;
            string teste = "";

            if (!System.IO.File.Exists(filepath))
            {
                return "Erro";
            }

            //Rafael - Abre o excel na 1ª sheet
            var excelDoc = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(filepath));
            var xlSht = excelDoc.Workbook.Worksheets.FirstOrDefault();
            int LinhaFinal=1;
            //'Conta linhas do ficheiro de excel
            linhaFinal = xlSht.Dimension.End.Row;

            string value1="sem valor";

            for (int i = 1; i < linhaFinal; i++)
            {
                value1=xlSht.GetValue(1, i).ToString();
                //'For n = 1 To 5000
                //if (xlSht.GetValue(i + 1, 1).ToString() != "")
                //{
                //    artigo = xlSht.GetValue(i + 1, 1).ToString();
                //}
            }
            return "";
        }
    }
}
