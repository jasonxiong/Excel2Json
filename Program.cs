using System;
using System.IO;
using System.Data;
using System.Text;
using Excel;
using System.Runtime.Remoting.Contexts;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace excel2json
{
    /// <summary>
    /// 应用程序
    /// </summary>
    sealed partial class Program
    {
        /// <summary>
        /// 应用程序入口
        /// </summary>
        /// <param name="args">命令行参数</param>
        static void Main(string[] args)
        {
            System.DateTime startTime = System.DateTime.Now;

            //-- 分析命令行参数
            var options = new Options();
            var parser = new CommandLine.Parser(with => with.HelpWriter = Console.Error);

            if (parser.ParseArgumentsStrict(args, options, () => Environment.Exit(-1)))
            {
                //-- 执行导出操作
                try
                {
                    Run(options);
                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error: " + exp.Message);
                }
            }

            //-- 程序计时
            System.DateTime endTime = System.DateTime.Now;
            System.TimeSpan dur = endTime - startTime;
            Console.WriteLine(
                string.Format("[{0}]：\t转换完成[{1}毫秒].",
                Path.GetFileName(options.ExcelPath),
                dur.Milliseconds)
                );
        }

        /// <summary>
        /// 根据命令行参数，执行Excel数据导出工作
        /// </summary>
        /// <param name="options">命令行参数</param>
        private static void Run(Options options)
        {
            string excelPath = options.ExcelPath;
            int header = options.HeaderRows;

            // 加载Excel文件
            DirectoryInfo excelFolder = new DirectoryInfo(excelPath);

            //加载json格式的导出文件列表
            string text = System.IO.File.ReadAllText(excelPath+ @"\FileList.json");
            JsonReader reader = new JsonTextReader(new StringReader(text));

            HashSet<String> FileListInfo = new HashSet<string> { };
            while (reader.Read())
            {
                if (reader.TokenType.ToString().CompareTo("String") != 0)
                {
                    continue;
                }

                //文件列表，读取到列表中
                FileListInfo.Add(reader.Value.ToString());
            }

            //遍历文件
            foreach (FileInfo configFile in excelFolder.GetFiles())
            {
                if(configFile.Extension != ".xlsx" && configFile.Extension!=".xls")
                {
                    //只处理表格
                    continue;
                }

                using (FileStream excelFile = configFile.OpenRead())    // File.Open(configFile., FileMode.Open, FileAccess.Read))
                {
                    // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelFile);

                    // The result of each spreadsheet will be created in the result.Tables
                    excelReader.IsFirstRowAsColumnNames = true;
                    DataSet book = excelReader.AsDataSet();

                    // 数据检测
                    if (book.Tables.Count < 1)
                    {
                        throw new Exception("Excel文件中没有找到Sheet: " + excelPath);
                    }

                    // 取得数据
                    for (int i = 0; i < book.Tables.Count; ++i)
                    {
                        if(!FileListInfo.Contains(book.Tables[i].TableName))
                        {
                            continue;
                        }

                        DataTable sheet = book.Tables[i];
                        if (sheet.Rows.Count <= 0)
                        {
                            //跳过
                            continue;
                            //throw new Exception("Excel Sheet中没有数据: " + excelPath);
                        }

                        //-- 确定编码
                        Encoding cd = new UTF8Encoding(false);
                        if (options.Encoding != "utf8-nobom")
                        {
                            foreach (EncodingInfo ei in Encoding.GetEncodings())
                            {
                                Encoding e = ei.GetEncoding();
                                if (e.EncodingName == options.Encoding)
                                {
                                    cd = e;
                                    break;
                                }
                            }
                        }

                        //-- 导出JSON文件
                        if (options.JsonPath != null && options.JsonPath.Length > 0)
                        {
                            JsonExporter exporter = new JsonExporter(sheet, header, options.Lowcase);
                            exporter.SaveToFile(options.JsonPath + book.Tables[i].TableName + ".json", cd);
                        }
                    }

                    //暂时不需要
                    /*
                    //-- 导出SQL文件
                    if (options.SQLPath != null && options.SQLPath.Length > 0)
                    {
                        SQLExporter exporter = new SQLExporter(sheet, header);
                        exporter.SaveToFile(options.SQLPath, cd);
                    }

                    //-- 生成C#定义文件
                    if (options.CSharpPath != null && options.CSharpPath.Length > 0)
                    {
                        string excelName = Path.GetFileName(excelPath);

                        CSDefineGenerator exporter = new CSDefineGenerator(sheet);
                        exporter.ClassComment = string.Format("// Generate From {0}", excelName);
                        exporter.SaveToFile(options.CSharpPath, cd);
                    }
                    */
                }
            }
        }
    }
}
