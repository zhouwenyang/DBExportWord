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
        static string _dataBaseName = "mintcode_tuotuo";
        static string _MySqlGetConnectionString
        {
            get
            {
                //return $"server=127.0.0.1;database={_dataBaseName};uid=root;pwd=perfect2018;";
                return $"server=172.16.8.7;database={_dataBaseName};uid=root;pwd=p@ssw0rd;";
            }
        }
        static string _MysqlGetTableSQL
        {
            get
            {
                return $"select table_name as tableName, table_comment as tableComment from information_schema.tables " +
                    $"where table_schema = '{_dataBaseName}' and table_type = 'base table';";
            }
        }

        static string MysqlGetTableSchemaSQL(string tableName)
        {
            return $"select " +
               $"COLUMN_NAME as code," +
               $"is_Nullable as IsNullable," +
               $"data_type AS datatype," +
               $"column_key as columnKey," +
               $"column_comment as comment," +
               $"CHARACTER_MAXIMUM_LENGTH as DataLength," +
               $"COLUMN_DEFAULT as defaultValue" +
               $" from information_schema.columns " +
               $"where table_schema = '{_dataBaseName}' and table_name = '{tableName}';";
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

            Console.WriteLine($"开始检索数据库{_dataBaseName}所有表...");
            var result = GetTableSchema<Table>(_MysqlGetTableSQL);

            Console.WriteLine($"总计 {result.Count} 表, 开始生成文档...");
            List<TableMapping> listMapping = new List<TableMapping>();
            listMapping.Add(new TableMapping() { DBColumn = "", DocumentColumn = "序号", ColumnWidth = 1000 });
            listMapping.Add(new TableMapping() { DBColumn = "code", DocumentColumn = "名称", ColumnWidth = 1000 });
            listMapping.Add(new TableMapping() { DBColumn = "comment", DocumentColumn = "描述", ColumnWidth = 2600, IsCenter = false });
            listMapping.Add(new TableMapping() { DBColumn = "DataType", DocumentColumn = "类型", ColumnWidth = 1000 });
            listMapping.Add(new TableMapping() { DBColumn = "DataLength", DocumentColumn = "长度", ColumnWidth = 1000 });
            listMapping.Add(new TableMapping() { DBColumn = "columnKey", DocumentColumn = "主键", ColumnWidth = 1000 });
            listMapping.Add(new TableMapping() { DBColumn = "IsNullable", DocumentColumn = "可空", ColumnWidth = 1000 });
            listMapping.Add(new TableMapping() { DBColumn = "defaultValue", DocumentColumn = "缺省", ColumnWidth = 1000 });
            ExportDocument(result, listMapping, $"{_dataBaseName}.docx");
            // select table_name from information_schema.tables where table_schema = 'csdb' and table_type = 'base table';

            Console.WriteLine($"数据库文档 {_dataBaseName} 生成成功");

            Console.Read();
        }



        /// <summary>
        /// 创建文档
        /// </summary>
        /// <param name="setting"></param>
        public static void ExportDocument(List<Table> listTable, List<TableMapping> listMapping, string fileName)
        {

            XWPFDocument doc = new XWPFDocument(NPOI.OpenXml4Net.OPC.OPCPackage.Open("12345.docx"));

            MemoryStream ms = new MemoryStream();
            DocumentSetting setting = new DocumentSetting();

            int docItemsCount = doc.BodyElements.Count;
            for (int i = 0; i < docItemsCount; i++)
            {
                doc.RemoveBodyElement(0);
            }

            //设置文档
            doc.Document.body.sectPr = new CT_SectPr();
            CT_SectPr setPr = doc.Document.body.sectPr;
            //获取页面大小
            Tuple<int, int> size = GetPaperSize(setting.PaperType);
            setPr.pgSz.w = (ulong)size.Item1;
            setPr.pgSz.h = (ulong)size.Item2;

            int tableNum = 0;
            foreach (var table in listTable)
            {
                tableNum++;

                string tableTitle = table.TableName;
                if (!string.IsNullOrEmpty(table.TableComment))
                {
                    tableTitle = $"{table.TableComment} {tableTitle}";
                }

                Console.WriteLine($"正在处理 {tableNum}.{tableTitle}");

                //创建一个段落
                XWPFParagraph gp = doc.CreateParagraph();

                XWPFStyles styles = doc.CreateStyles();

                //表格段落标题
                gp.SetNumID($"2.1");
                gp.Style = "Heading2";
                XWPFRun gr = gp.CreateRun();
                gr.SetText($"{tableTitle}");

                //获取表字段信息
                var listTableSchema = GetTableSchema<TableSchema>(MysqlGetTableSchemaSQL(table.TableName));

                //创建表格
                XWPFTable docTable = doc.CreateTable(listTableSchema.Count + 1, listMapping.Count);
                int i = 0;
                foreach (var tableMapping in listMapping)
                {
                    var currentCell = docTable.GetRow(0).GetCell(i);
                    currentCell = SetCell(currentCell, tableMapping.DocumentColumn, tableMapping.ColumnWidth, "CCCCCC", true, 10, true);
                    //var ppr = currentCell.GetCTTc().AddNewP().AddNewPPr();
                    //space.line="0";
                    //space.lineRule = ST_LineSpacingRule.auto;
                    i++;
                }
                int rowIndex = 1;
                foreach (var tableSchema in listTableSchema)
                {
                    int cellIndex = 0;
                    foreach (var tableMapping in listMapping)
                    {
                        var currentCell = docTable.GetRow(rowIndex).GetCell(cellIndex);
                        string color = string.Empty;
                        string text = string.Empty;
                        if (rowIndex % 2 == 0)
                        {
                            color = "DDDDDD";
                        }

                        if (string.IsNullOrEmpty(tableMapping.DBColumn))
                        {
                            text = rowIndex.ToString();
                        }
                        else
                        {
                            text = GetPropValue(tableSchema, tableMapping.DBColumn);
                        }

                        currentCell = SetCell(currentCell, text, tableMapping.ColumnWidth, color, tableMapping.IsCenter, 10);

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

        private static XWPFTableCell SetCell(XWPFTableCell cell, string text, int width, string color = "", bool isCenter = true, int fontSize = 0, bool isBold = false)
        {

            CT_Tc cttc = cell.GetCTTc();
            CT_TcPr ctpr = cttc.AddNewTcPr();
            //垂直居中
            ctpr.AddNewVAlign().val = ST_VerticalJc.center;
            //设置单元格列宽
            ctpr.tcW = new CT_TblWidth();
            ctpr.tcW.w = width.ToString();
            ctpr.tcW.type = ST_TblWidth.dxa;
            //设置单元格背景色
            if (!string.IsNullOrEmpty(color))
            {
                cell.SetColor(color);
            }
            //设置文本
            text = text == "PRI" ? "Y" : text;
            text = text == "YES" ? "Y" : text;
            text = text == "NO" ? "N" : text;
            text = text == "CURRENT_TIMESTAMP" ? "now()" : text;
            cell.RemoveParagraph(0);
            XWPFParagraph gp = cell.AddParagraph();
            if (isCenter)
            {
                //水平居中
                gp.Alignment = ParagraphAlignment.CENTER;
            }
            //文字格式
            gp.Style = "NoSpacing";
            XWPFRun gr = gp.CreateRun();
            if (fontSize > 0)
            {
                gr.FontSize = fontSize;
            }
            gr.IsBold = isBold;
            gr.SetText(text);

            return cell;
        }

        /// <summary>
        /// 根据字段获取属性的值
        /// </summary>
        /// <param name="src"></param>
        /// <param name="propName"></param>
        /// <returns></returns>
        private static string GetPropValue(object data, string propName)
        {
            object value = data.GetType().GetProperty(propName).GetValue(data, null);
            return null == value ? string.Empty : value.ToString();
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
        public string DocumentColumn { get; set; }
        public int ColumnWidth { get; set; }

        public bool IsCenter { get; set; } = true;
    }
}
