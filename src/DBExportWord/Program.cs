using Dapper;
using MySql.Data.MySqlClient;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static ExportWordUtils;

namespace DBExportWord
{
    class Program
    {
        static string _MySqlGetConnectionString
        {
            get
            {
                return "server=172.16.8.7;database=mintcode_tuotuo;uid=root;pwd=p@ssw0rd;";
            }
        }
        static string _MysqlGetTableSQL
        {
            get
            {
                return "select table_name as tableName, table_comment as tableComment from information_schema.tables where table_schema = 'mintcode_tuotuo' and table_type = 'base table';";
            }
        }

        static string MysqlGetTableSchemaSQL(string tableName)
        {
            return $"select " +
               $"COLUMN_NAME as code,is_Nullable as IsNullable,data_type AS datatype,column_key as columnKey,column_comment as comment,CHARACTER_MAXIMUM_LENGTH as DataLength,COLUMN_DEFAULT as defaultValue" +
               $" from information_schema.columns where table_schema = 'mintcode_tuotuo' and table_name = '{tableName}';";
        }

        public static List<T> GetTableSchema<T>(string sql)
        {
            using (var mysql = new MySqlConnection(_MySqlGetConnectionString))
            {

                mysql.Open();
                var result = mysql.Query<T>(sql);

                return result.AsList();
            }
        }

        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = Encoding.GetEncoding("GB2312");

            Console.WriteLine("开始检索所有表...");
            var result = GetTableSchema<Table>(_MysqlGetTableSQL);

            Console.WriteLine($"总计 {result.Count} 表, 开始生成文档...");
            List<TableMapping> listMapping = new List<TableMapping>();
            listMapping.Add(new TableMapping() { DBColumn = "", DocumentColum = "序号" });
            listMapping.Add(new TableMapping() { DBColumn = "code", DocumentColum = "名称" });
            listMapping.Add(new TableMapping() { DBColumn = "comment", DocumentColum = "描述" });
            listMapping.Add(new TableMapping() { DBColumn = "DataType", DocumentColum = "类型" });
            listMapping.Add(new TableMapping() { DBColumn = "DataLength", DocumentColum = "长度" });
            listMapping.Add(new TableMapping() { DBColumn = "columnKey", DocumentColum = "主键" });
            listMapping.Add(new TableMapping() { DBColumn = "IsNullable", DocumentColum = "可空" });
            listMapping.Add(new TableMapping() { DBColumn = "defaultValue", DocumentColum = "缺省值" });

            ExportDocument(result, listMapping, "123.doc");
            // select table_name from information_schema.tables where table_schema = 'csdb' and table_type = 'base table';

            Console.WriteLine($"文档生成成功");

            Console.Read();
        }



        /// <summary>
        /// 创建文档
        /// </summary>
        /// <param name="setting"></param>
        public static void ExportDocument(List<Table> listTable, List<TableMapping> listMapping, string fileName)
        {
            XWPFDocument doc = new XWPFDocument();
            MemoryStream ms = new MemoryStream();
            DocumentSetting setting = new DocumentSetting();

            //设置文档
            doc.Document.body.sectPr = new CT_SectPr();
            CT_SectPr setPr = doc.Document.body.sectPr;
            //获取页面大小
            Tuple<int, int> size = GetPaperSize(setting.PaperType);
            setPr.pgSz.w = (ulong)size.Item1;
            setPr.pgSz.h = (ulong)size.Item2;
            //创建一个段落
            CT_P p = doc.Document.body.AddNewP();
            //段落水平居中
            p.AddNewPPr().AddNewJc().val = ST_Jc.center;
            XWPFParagraph gp = new XWPFParagraph(p, doc);

            XWPFRun gr = gp.CreateRun();
            //创建标题
            if (!string.IsNullOrEmpty(setting.TitleSetting.Title))
            {
                gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = setting.TitleSetting.FontName;
                gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = setting.TitleSetting.FontName;
                gr.GetCTR().AddNewRPr().AddNewRFonts().hint = ST_Hint.eastAsia;
                gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)setting.TitleSetting.FontSize;//2号字体
                gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)setting.TitleSetting.FontSize;
                gr.GetCTR().AddNewRPr().AddNewB().val = setting.TitleSetting.HasBold; //加粗
                gr.GetCTR().AddNewRPr().AddNewColor().val = "black";//字体颜色
                gr.SetText(setting.TitleSetting.Title);
            }
            int tableNum = 0;
            foreach (var table in listTable)
            {
                tableNum++;
                Console.WriteLine($"正在处理表 {tableNum}.{table.TableName} 表");


                p = doc.Document.body.AddNewP();
                p.AddNewPPr().AddNewJc().val = ST_Jc.both;
                gp = new XWPFParagraph(p, doc)
                {
                    //IndentationFirstLine = 2
                };
                gp.Style = "Heading1";
                gp.SetNumID($"1.1.{tableNum}");
                //单倍为默认值（240）不需设置，1.5倍=240X1.5=360，2倍=240X2=480
                //p.AddNewPPr().AddNewSpacing().line = "400";//固定20磅
                //p.AddNewPPr().AddNewSpacing().lineRule = ST_LineSpacingRule.exact;

                gr = gp.CreateRun();
                //CT_RPr rpr = gr.GetCTR().AddNewRPr();
                //CT_Fonts rfonts = rpr.AddNewRFonts();
                //rfonts.ascii = setting.MainContentSetting.FontName;
                //rfonts.eastAsia = setting.MainContentSetting.FontName;
                //rpr.AddNewSz().val = (ulong)setting.MainContentSetting.FontSize;//5号字体-21
                //rpr.AddNewSzCs().val = (ulong)setting.MainContentSetting.FontSize;
                //rpr.AddNewB().val = setting.MainContentSetting.HasBold;
                gr.SetText($"{table.TableName}({table.TableComment})");

                var listTableSchema = GetTableSchema<TableSchema>(MysqlGetTableSchemaSQL(table.TableName));

                XWPFTable docTable = doc.CreateTable(listTableSchema.Count + 1, 8);

                int i = 0;
                foreach (var tableHeader in listMapping)
                {
                    docTable.GetRow(0).GetCell(i).SetColor("#CCCCCC");
                    XWPFParagraph pIO = docTable.GetRow(0).GetCell(i).AddParagraph();
                    pIO.Alignment = ParagraphAlignment.CENTER;

                    XWPFRun rIO = pIO.CreateRun();
                    rIO.FontSize = 10;
                    rIO.IsBold = true;
                    rIO.SetText(tableHeader.DocumentColum);
                    i++;
                }
                int rowIndex = 1;
                foreach (var tableSchema in listTableSchema)
                {
                    int cellIndex = 0;
                    foreach (var tableHeader in listMapping)
                    {
                        var currentCell = docTable.GetRow(rowIndex).GetCell(cellIndex);
                        if (rowIndex % 2 == 0)
                        {
                            currentCell.SetColor("#DDDDDD");
                        }

                        XWPFParagraph pIO = currentCell.AddParagraph();
                        XWPFRun rIO = pIO.CreateRun();
                        rIO.FontSize = 9;
                        if (string.IsNullOrEmpty(tableHeader.DBColumn))
                        {
                            rIO.SetText(rowIndex.ToString());
                            pIO.Alignment = ParagraphAlignment.CENTER;
                        }
                        else
                        {
                            rIO.SetText(tableHeader.DBColumn);
                        }
                        cellIndex++;
                    }
                    rowIndex++;
                }

            }

            //开始写入
            doc.Write(ms);

            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
            }
            ms.Close();
        }
    }

    public class Table
    {
        public string TableName { get; set; }
        public string TableComment { get; set; }

    }

    public class TableSchema
    {
        public string code { get; set; }
        public string comment { get; set; }
        public string DataType { get; set; }
        public string DataLength { get; set; }
        public string columnKey { get; set; }
        public string IsNullable { get; set; }
        public string defaultValue { get; set; }
    }


    /// <summary>
    /// 映射
    /// </summary>
    public class TableMapping
    {
        public string DBColumn { get; set; }
        public string DocumentColum { get; set; }

    }
}
