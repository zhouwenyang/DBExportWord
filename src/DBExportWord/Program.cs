using Dapper;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Text;
using static ExportWordUtils;

namespace DBExportWord
{
    class Program
    {
        static string _connectionString = "server=172.16.8.7;database=mintcode_tuotuo;uid=root;pwd=p@ssw0rd;";

        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.GetEncoding("GB2312");

            using (var mysql = new MySqlConnection(_connectionString))
            {
                mysql.Open();
                var result = mysql.Query("select table_name, table_comment from information_schema.tables where table_schema = 'mintcode_tuotuo' and table_type = 'base table';");

                string content = string.Empty;
                foreach(var table in result)
                {
                    IDictionary<string, object> dict = table;


                    content += $"{dict["table_name"]} {dict["table_comment"]}";


                }
                ExportWord(content);

            }

            // select table_name from information_schema.tables where table_schema = 'csdb' and table_type = 'base table';


            Console.WriteLine("Hello World!");
            Console.Read();
        }

        private static void ExportToWord(string content)
        {
            DocumentSetting setting = new DocumentSetting();
            setting.MainContentSetting.MainContent = content;
            setting.TitleSetting.Title = "测试";
            setting.SavePath = "12345.doc";
            ExportWordUtils.ExportDocument(setting);
        }
    }
}
