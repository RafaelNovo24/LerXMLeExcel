using System;
using System.Windows.Forms;
using ManipularExcel;

namespace ManipularExcel
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("**************** Bem vindo ao menu de testes Excel ****************");
            Console.WriteLine("Escolha a opção:");
            Console.WriteLine($@"1.-Importar Excel.");
            Console.WriteLine($@"2.-Sair");

            string resposta = Console.ReadLine();

            switch (resposta)
            {
                case "1":
                    ImportarExcel();
                    
                    break;

                case "2":
                    
                    break;
                default:
                    break;
            }

            [STAThread]
            static string ImportarExcel()
            {
                string filePath="";
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
        }
    }
}
