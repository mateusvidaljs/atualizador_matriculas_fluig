using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using System.Runtime.InteropServices;
using System.Net;
using AtualizadorMatriculas.FluigColleagueService;
using System.Xml;

namespace AtualizadorMatriculas
{
    class Program
    {
        static void Main(string[] args)
        {
            Program classe = new Program();
            classe.init();
        }

        private string formatNewUserRegistre(string matricula)
        {
            int tamatricula = matricula.Length;
            int i;

            if(matricula.Length < 6 && matricula.Substring(0,1) == "0")
            {
                for(i = tamatricula; i < 6; i++)
                {
                    matricula = "0" + matricula;
                }
            }

            return matricula;
        }

        public bool copyAllIds(string matricula)
        {
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "VM30797";
                builder.UserID = "sa";
                builder.Password = "jugu1ch@cdmg";
                builder.InitialCatalog = "fluig";

                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    string sql = "UPDATE FDN_USERDATA SET DATA_VALUE = '" + matricula + "' FROM FDN_USERDATA D JOIN FDN_USERTENANT B ON D.USER_TENANT_ID = B.USER_TENANT_ID WHERE B.USER_CODE = '" + matricula + "' AND D.DATA_KEY = 'UserProjects'";
                    //connection.Open();
                    
                    SqlCommand command = new SqlCommand(sql, connection);
                    connection.Open();
                    command.ExecuteNonQuery();
                }

                return true;
            }
            catch (SqlException ex)
            {
                Console.WriteLine(ex);
                return false;
            }
        }

        public void init()
        {
            string menuoption = "";
            /** CONFIGURAÇÕES GERAIS DO APP */
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
            Console.WriteLine("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx ATUALIZADOR DE MATRICULAS xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
            Console.WriteLine("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\n");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("c::::> @author: Mateus Vidal");
            Console.WriteLine("c::::> @enterprise: WebMoon Design MEI");
            Console.WriteLine("c::::> @mail: mateusvidal.dev@gmail.com");
            Console.WriteLine("c::::> @costumer: CODEMGE - GETIN\n");
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.DarkGreen;

            try
            {
                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.Yellow;

                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                builder.DataSource = "VM30797";
                builder.UserID = "sa";
                builder.Password = "jugu1ch@cdmg";
                builder.InitialCatalog = "fluig";

                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    StringBuilder sb = new StringBuilder();
                    sb.Append("SELECT COUNT(*) AS 'USUARIOS SEM PROJETO' FROM FDN_USERDATA WHERE DATA_VALUE = '' AND DATA_KEY = 'UserProjects'");
                    String sql = sb.ToString();

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Console.WriteLine("Existem, atualmente, " + reader.GetInt32(0).ToString() + " usuários sem ID PROJETO\n");
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                throw ex;
            }

            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.DarkGreen;

            Console.WriteLine("Menu (Digite a opção desejada):");
            Console.WriteLine("[1] - Atualizar Matrículas");
            Console.WriteLine("[2] - Listar Usuários Sem ID Projeto");
            Console.WriteLine("[3] - Copiar Matrículas");
            Console.WriteLine("Digite a opção...");

            //Recebe e interpreta a variavel digitada
            menuoption = Console.ReadLine().ToString();
                        
            if(menuoption == "1")
            {
                Console.Clear();

                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("[ x=x=x=x=x=x=x=x=x EXECUTANDO ATUALIZAÇÃO  x=x=x=x=x=x=x=x=x ]");

                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.White;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath("C:/Projetos/DotNet/Conversao.xlsx"));
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);
                Excel.Range xlRange = xlWorksheet.UsedRange;
                object[,] valueArray = (object[,])xlRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                Console.WriteLine("+-----------------+");
                Console.WriteLine("| Atual  |  Nova  |");

                for (int row = 1; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= xlWorksheet.UsedRange.Columns.Count; ++col)
                    {
                        Console.WriteLine("+-----------------+");
                        Console.WriteLine("| " + valueArray[row, 1].ToString() + " | " + valueArray[row, 2].ToString() + " |");
                        break;
                    }
                }
                
                xlWorkbook.Close(false);
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);

                Console.WriteLine("Deseja prosseguir com a atualização? (S/N)");
                Console.WriteLine("Digite a opção...");
                string confirmupdate = Console.ReadLine().ToString();

                if(confirmupdate == "s" || confirmupdate == "S")
                {

                }
                else
                {
                    
                }
            }
            else if (menuoption == "2")
            {
                try
                {
                    SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
                    builder.DataSource = "VM30797";
                    builder.UserID = "sa";
                    builder.Password = "jugu1ch@cdmg";
                    builder.InitialCatalog = "fluig";

                    using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                    {
                        connection.Open();
                        StringBuilder sb = new StringBuilder();
                        sb.Append("SELECT T.DATA_KEY, T.DATA_VALUE, T.USER_TENANT_ID, D.USER_CODE, D.EMAIL, D.LOGIN FROM FDN_USERDATA AS T JOIN FDN_USERTENANT AS D ON T.USER_TENANT_ID = D.USER_TENANT_ID WHERE T.DATA_KEY = 'UserProjects' AND T.DATA_VALUE = ''");
                        String sql = sb.ToString();

                        using (SqlCommand command = new SqlCommand(sql, connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    Console.WriteLine("+-----------------------------------------------+");
                                    Console.WriteLine("|  USUÁRIO: " + reader.GetString(5));
                                }
                            }
                        }
                    }
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
            }
            else if(menuoption == "3")
            {
                Console.Clear();

                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("[ x=x=x=x=x=x=x=x=x EXECUTANDO CÓPIA DE MATRÍCULAS x=x=x=x=x=x=x=x=x ]");

                Console.ResetColor();
                Console.ForegroundColor = ConsoleColor.White;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath("C:/Projetos/DotNet/Conversao.xlsx"));
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);
                Excel.Range xlRange = xlWorksheet.UsedRange;
                object[,] valueArray = (object[,])xlRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

                Console.WriteLine("+-----------------+");
                Console.WriteLine("| Atual  |  Nova  |");

                for (int row = 1; row <= xlWorksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= xlWorksheet.UsedRange.Columns.Count; ++col)
                    {
                        this.copyAllIds(valueArray[row, 1].ToString());
                        //copyAllIds(string matricula);
                        Console.WriteLine("+-----------------+");
                        Console.WriteLine("| " + valueArray[row, 1].ToString() + " | " + valueArray[row, 2].ToString() + " | ::::::> EXECUTADO!");
                        break;
                    }
                }

                xlWorkbook.Close(false);
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
            }

            Console.ReadKey();
        }
    }
}
