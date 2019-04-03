using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;



namespace excel_json
{
    class Program
    {
        static void Main(string[] args)
        {
            var xml = "<t><machNumber>4</machNumber>" +
          "<machNumber>5</machNumber>" +
          "<machNumber>13</machNumber>" +
          "<machNumber>14</machNumber>" +
          "<machNumber>15</machNumber>" +
          "<machNumber>45</machNumber>" +
          "<machNumber>46</machNumber>" +
          "<machNumber>18</machNumber>" +
          "<machNumber>6</machNumber>" +
          "<machNumber>8</machNumber>" +
          "<machNumber>28</machNumber>" +
          "<machNumber>29</machNumber>" +
          "<machNumber>1</machNumber>" +
          "<machNumber>2</machNumber>" +
          "<machNumber>3</machNumber>" +
          "<machNumber>21</machNumber>" +
          "</t>";


            FileInfo filename = new FileInfo(@"d:\Отчет.xlsx");

            if (args.Length > 0)
            {
                Console.WriteLine("Обработка данных");


                foreach (Object arg in args)
                {
                    if (arg.Equals("-J"))
                    {
                        Console.WriteLine("Генерирую JSON файл");
                        GenJson(xml, null);
                    }

                    if (arg.Equals("-E"))
                    {
                        Console.WriteLine("Генерирую Excel файл");
                        GenExcel(xml, filename);
                    }
                }

            }
            else
            {
                Console.WriteLine("Генератор Json и Excel \n");
                Console.WriteLine("Использование: \n excel_json.exe [аргумент] [имя файла] ");
                Console.WriteLine("\t -J \t создать JSON файл");
                Console.WriteLine("\t -E \t создать EXCEL файл");
            }
        }


        private static void GenJson(object data, FileInfo filename)
        {

            //var smpoData = new List<JSMPOClass>();

            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();



            
            
                var temperatures = FromJson(jsonString);


            try
            {


                JSMPO jSMPO = new JSMPO {
                    
                    date = "Февраль 2019",
                    unit = "Часов",
                    columns = new List<Columns>(2)
                };





                //jSMPO.date = "Февраль 2019";
                //jSMPO.unit = "Часов";

                //Datum dataj = new Datum();
                //dataj.item = "DORRIES поз. 51 - 9";
                //dataj.Вналадке = "101";
                //dataj.Впростое = "57";
                //dataj.Вработе = "514";


                //Columns col = new Columns();
                //col.column = "В работе";
                //col.color = "#9BBB59";



                







                //smpo.Add(new Datum(new Data("DORRIES поз. 51-9", "101", "57", "514"), "Февраль 2019", "Часов", new Columns("item", "#9BBB59")));
                //smpo.Add(new Datum(new Data("DORRIES поз. 51-9", "101", "57", "514"), "Февраль 2019", "Часов", new Columns("item", "#9BBB59")));
                //smpo.Add(new Datum(new Data("DORRIES поз. 51-9", "101", "57", "514"), "Февраль 2019", "Часов", new Columns("item", "#9BBB59")));
                //smpo.Add(new Datum(new Data("DORRIES поз. 51-9", "101", "57", "514"), "Февраль 2019", "Часов", new Columns("item", "#9BBB59")));
                //smpo.Add(new Datum(new Data("DORRIES поз. 51-9", "101", "57", "514"), "Февраль 2019", "Часов", new Columns("item", "#9BBB59")));

                // JSMPOData smpo = new JSMPOData(new Data("DORRIES поз. 51-9", "101", "57", "514"), "Февраль 2019", "Часов", new Columns("item", "#9BBB59"));
                //JSMPOData smpo2 = new JSMPOData(new Data("FA-C221 поз. 51-5", "215", "112", "345"), "Февраль 2019", "Часов", new Columns());


                //   JSMPOData[] smpo = new JSMPOData[] { smpo1};

                string json = JsonConvert.SerializeObject(jSMPO, Formatting.Indented);


                Console.WriteLine(json);


                /*  smpo.data.Add("DORRIES поз. 51-9");
                  smpo.data.Add("FA-C221 поз. 51-5");
                  smpo.data.Add("FA-C221 поз. 51-5");
                  smpo.data.Add("FA-C221 поз. 51-5");*/


                JsonSerializer serializer = new JsonSerializer();
                serializer.Converters.Add(new JavaScriptDateTimeConverter());
                serializer.NullValueHandling = NullValueHandling.Ignore;

                using (StreamWriter file = new StreamWriter(@"d:\smpo.json"))
                using (JsonWriter writer = new JsonTextWriter(file))
                {
                    serializer.Serialize(writer, jSMPO);
                    // {"ExpiryDate":new Date(1230375600000),"Price":0}
                }


            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally
            {
                // Закрыть соединение.
                conn.Close();
                // Разрушить объект, освободить ресурс.
                conn.Dispose();
            }

        }


        private static void GenJson_old(object data, FileInfo filename)
        {
            int x_diag = 10, y_diag = 10;
            var smpoData = new List<SMPO>();

            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();

            try
            {
                using (var exclelFile = new ExcelPackage(filename))
                {
                    smpoData = QueryEmployee(conn);

                    // Добавим новые листы
                    var gist = exclelFile.Workbook.Worksheets.Add("Графики");
                    var ws = exclelFile.Workbook.Worksheets.Add("Данные");


                    // Создаем таблицу на странице Сводная таблица
                    var row = 1;
                    var col = 1;

                    var tempdate = "";

                    foreach (var item in smpoData)
                    {
                        if (!item.date_start.Equals(tempdate))
                        {
                            ws.Cells[row, col].Value = item.date_start;

                            // Создаем гистограмму

                            var chart = (ExcelBarChart)gist.Drawings.AddChart(item.date_start, eChartType.ColumnClustered);
                            chart.SetSize(480, 320);
                            chart.SetPosition(x_diag, y_diag);
                            chart.Title.Text = string.Format("Гистограмма работы станков от {0}", item.date_start);

                            chart.Series.Add(ExcelRange.GetAddress(3, 2, 4, 2), ExcelRange.GetAddress(3, 1, 4, 1));


                            // Настраиваем шапку даты
                            using (var range = ws.Cells[row, col, row, 7])
                            {
                                range.Merge = true;
                                range.Style.Font.Bold = true;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                range.Style.Font.Color.SetColor(Color.White);
                            }

                            row++;

                            ws.Cells[row, col++].Value = "Номер станка";
                            ws.Cells[row, col++].Value = "Имя станка";
                            ws.Cells[row, col++].Value = "в работе";
                            ws.Cells[row, col++].Value = "в наладке";
                            ws.Cells[row, col++].Value = "в простое";
                            ws.Cells[row, col++].Value = "кол-во часов в месяце";
                            ws.Cells[row, col++].Value = "% работы станка";
                            row++;
                            col = 1;
                            x_diag = x_diag + 340;
                        }
                        ws.Cells[row, col++].Value = item.IdStanok;
                        ws.Cells[row, col++].Value = item.name_stan;



                        gist.Cells[string.Format("A{0}", row)].Value = item.name_stan; // Гист
                        gist.Cells[string.Format("B{0}", row)].Value = Convert.ToInt32(item.ВРаботе);

                        ws.Cells[row, col++].Value = Convert.ToInt32(item.ВРаботе);
                        ws.Cells[row, col++].Value = Convert.ToInt32(item.ВНаладке);
                        ws.Cells[row, col++].Value = Convert.ToInt32(item.ВПростое);
                        ws.Cells[row, col++].Value = Convert.ToInt32(item.КолвоЧасовВМесяце);
                        ws.Cells[row, col++].Value = Convert.ToInt32(item.ПроцРаботыСтанка);


                        row++;
                        col = 1;
                        tempdate = item.date_start;
                    }

                    // добавим всем ячейкам рамку

                    using (var cells = ws.Cells[ws.Dimension.Address])
                    {
                        cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        cells.AutoFitColumns();
                    }

                    exclelFile.Save();

                    // сохраняем в файл
                    //var bin = p.GetAsByteArray();
                    //File.WriteAllBytes(@"d:\result.xlsx", bin);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally
            {
                // Закрыть соединение.
                conn.Close();
                // Разрушить объект, освободить ресурс.
                conn.Dispose();
            }

        }



        private static void GenExcel(object data, FileInfo filename)
        {
            // Вызов процедуры
            var tempdate = "";
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();

            try
            {
                using (var exclelFile = new ExcelPackage(filename))
                {
                    var row = 1;
                    var col = 1;


                    // Добавим новые листы
                    //  var gist = exclelFile.Workbook.Worksheets.Add("Графики");
                    //  var ws = exclelFile.Workbook.Worksheets.Add("Данные");


                    ExcelWorksheet gist = exclelFile.Workbook.Worksheets["Графики"];
                    ExcelWorksheet ws = exclelFile.Workbook.Worksheets["Данные"];


                    // Устанавливаем крайнии координаты гистограмме
                    var maxPosRow = 0;
                    //Ищем крайнюю гистограмму на листе
                    for (int i = 0; i < gist.Drawings.Count; i++)
                    {
                        if (gist.Drawings[i].To.Row > maxPosRow)
                        {
                            maxPosRow = gist.Drawings[i].To.Row + 2;
                        };
                    }

                    // 20px
                    int x_diag = maxPosRow * 20, y_diag = 0;

                    // Ищем конец таблицы
                    row = ws.Dimension.Rows+1;                   



                    // Создать объект Command для вызова процедуры Get_Employee_Info.
                    SqlCommand cmd = new SqlCommand("reports.machine_performance", conn);

                    // Вид Command является StoredProcedure
                    cmd.CommandType = CommandType.StoredProcedure;


                    /* Выборка за период*/

                    var startOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                    //var endOfPreviousMonth = startOfMonth.AddDays(-1);
                    var startOfPreviousMonth = new DateTime(startOfMonth.AddDays(-1).Year, startOfMonth.AddDays(-1).Month, 1);
                    var endOfPreviousPreviousMonth = startOfMonth.AddMonths(0).AddDays(-1);


                    //cmd.Parameters.Add("@date_start", SqlDbType.Date).Value = startOfPreviousMonth;
                    //cmd.Parameters.Add("@date_end", SqlDbType.Date).Value = endOfPreviousPreviousMonth;

                    cmd.Parameters.Add("@date_start", SqlDbType.Date).Value = "01.03.2019";
                    cmd.Parameters.Add("@date_end", SqlDbType.Date).Value = "31.03.2019";

                    cmd.Parameters.Add("@machine_list", SqlDbType.Xml).Value = data;


                    // Выполнить процедуру.

                    // получаем список колонок таблицы
                    using (var reader = cmd.ExecuteReader())
                    {

                        if (reader.HasRows)
                        {
                            while (reader.Read()) // Пробегаемся по всем нашим записям
                            {
                                //  reader.Read();
                                var tableSchema = reader.GetSchemaTable();

                                if (!reader[0].Equals(tempdate)) // Если не таже дата то создаем новый заголовок
                                {
                                    // Настраиваем стиль шапки даты
                                    using (var range = ws.Cells[row, col, row, tableSchema.Rows.Count - 1])
                                    {
                                        range.Merge = true;
                                        //range.Style.Font.Bold = true;
                                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        //range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        //range.Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                        //range.Style.Font.Color.SetColor(Color.White);
                                    }

                                    ws.Cells[row, col].Value = reader[0]; // Создаем шапку с датой


                                    row++;

                                    foreach (DataRow row_name in tableSchema.Rows)
                                    {

                                        if (row_name["ColumnName"].Equals("Статус оборудования"))
                                        {
                                            //задаем стиль колонке
                                            ws.Cells[row, col].Style.Font.Bold = true;
                                        }

                                        if (!row_name["ColumnName"].Equals("Период"))
                                        {
                                            ws.Cells[row, col++].Value = row_name["ColumnName"];
                                        }
                                    }


                                    ExcelChart chart = (ExcelChart)gist.Drawings[0];
                                    ExcelChart cc = gist.Drawings.AddChart(Convert.ToString(reader[0]), eChartType.ColumnStacked);



                                    // Копируем предыдущую версию стилей
                                    var xml = XDocument.Parse(chart.ChartXml.InnerXml);
                                    XNamespace nsC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                                    XNamespace nsA = "http://schemas.openxmlformats.org/drawingml/2006/main";

                                    // Загружаем данные в гистограмму абсолютный путь
                                    var fs = xml.Descendants(nsC + "f");
                                    foreach (var f in fs)
                                    {
                                        f.Value = ws.Cells[f.Value].Offset(row - 2, 0).FullAddressAbsolute;
                                    }

                                    // Загружаем стиль из xml.
                                    cc.ChartXml.InnerXml = xml.ToString();


                                    // Создаем гистограмму
                                    // var chart = (ExcelBarChart)gist.Drawings.AddChart(Convert.ToString(reader[0]), eChartType.ColumnStacked);
                                    cc.SetSize(1169, 479);
                                    cc.SetPosition(x_diag, y_diag);
                                    cc.Title.Text = string.Format("{0}", reader[0]);

                                    //string serieAddress, xSerieAddress;

                                    //serieAddress = ExcelCellBase.GetFullAddress("Данные", ExcelCellBase.GetAddress(row + 1, 2, row + 1, tableSchema.Rows.Count - 1));
                                    //// chart.Legend.Border.Fill.Color = Color.Yellow; // Координаты данных

                                    //var headerAddr = ws.Cells[row + 1, 1];
                                    ////headerAddr.Style.Font.Bold = true;
                                    //xSerieAddress = ExcelCellBase.GetFullAddress("Данные", ExcelCellBase.GetAddress(row, 2, row, tableSchema.Rows.Count - 1)); //Координаты легедны
                                    //cc.Series.Add(serieAddress, xSerieAddress).HeaderAddress = headerAddr;
                                    ////chart.Legend.Border.Fill.Color = Color.Yellow;


                                    //headerAddr = ws.Cells[row + 2, 1];
                                    //serieAddress = ExcelCellBase.GetFullAddress("Данные", ExcelCellBase.GetAddress(row + 2, 2, row + 2, tableSchema.Rows.Count - 1)); // Координаты данных
                                    //cc.Series.Add(serieAddress, xSerieAddress).HeaderAddress = headerAddr;

                                    //headerAddr = ws.Cells[row + 3, 1];
                                    //serieAddress = ExcelCellBase.GetFullAddress("Данные", ExcelCellBase.GetAddress(row + 3, 2, row + 3, tableSchema.Rows.Count - 1)); // Координаты данных
                                    //cc.Series.Add(serieAddress, xSerieAddress).HeaderAddress = headerAddr;
                                    ////chart.Legend.Border.Fill.Color = Color.Yellow;


                                    // снова получаем крайнии координаты
                                    for (int i = 0; i < gist.Drawings.Count; i++)
                                    {
                                        if (gist.Drawings[i].To.Row > maxPosRow)
                                        {
                                            maxPosRow = gist.Drawings[i].To.Row + 2;
                                        };
                                    }

                                    x_diag = maxPosRow * 20;


                                    row++;
                                    col = 1;
                                }

                                ws.Cells[row, col++].Value = Convert.ToString(reader[1]);



                                if (Convert.ToString(reader[1]) == "Статус оборудования")
                                {
                                    ws.Cells[row, col - 1].Style.Font.Bold = true;
                                }


                                if (Convert.ToString(reader[1]) == "% работы станка")
                                {
                                    ws.Cells[row, col - 1].Style.Font.Bold = true;
                                    // Вычисляем процент
                                    for (var i = 2; i <= tableSchema.Rows.Count - 1; i++)
                                    {
                                        // Задаем стиль
                                        ws.Cells[row, col].Style.Font.Bold = true;
                                        ws.Cells[row, col].Style.Font.Color.SetColor(Color.Red);
                                        ws.Cells[row, col].Style.Numberformat.Format = "0%";
                                        ws.Cells[row, col++].Formula = string.Format("=(({0}+{1})/{2})", ws.Cells[row - 4, col - 1].Address,
                                            ws.Cells[row - 3, col - 1].Address, ws.Cells[row - 1, col - 1].Address);
                                    }
                                }
                                else
                                {
                                    // Заполняем таблицу значениями
                                    for (var i = 2; i <= tableSchema.Rows.Count - 1; i++)
                                    {
                                        // Проверка на пустные значания
                                        if (!DBNull.Value.Equals(reader[i]))
                                            ws.Cells[row, col++].Value = Convert.ToInt32(reader[i]);
                                        else
                                            ws.Cells[row, col++].Value = 0;
                                    }
                                }

                                row++;
                                col = 1;
                                tempdate = Convert.ToString(reader[0]);

                            }
                        }
                    }

                    // добавим всем ячейкам рамку
                    using (var cells = ws.Cells[ws.Dimension.Address])
                    {
                        cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        cells.AutoFitColumns();
                    }


                    // сохраняем в файл

                    //exclelFile.Save();
                    var bin = exclelFile.GetAsByteArray();
                    File.WriteAllBytes(Convert.ToString(filename), bin);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
        }


        private static List<SMPO> QueryEmployee(SqlConnection conn)
        {
            string sql = "SELECT [id_event],[event_name] FROM[SMPO].[dbo].[event_excel]";

            // Создать объект Command.
            SqlCommand cmd = new SqlCommand();

            // Сочетать Command с Connection.
            cmd.Connection = conn;
            cmd.CommandText = sql;


            List<SMPO> dataSmpos = new List<SMPO>();

            using (DbDataReader reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {

                    while (reader.Read())
                    {
                        var date_start = reader.GetValue(0);
                        var IdStanok = Convert.ToInt32(reader.GetString(1));
                        var name_stan = reader.GetString(2);
                        var вработе = reader.GetValue(3);
                        var вналадке = reader.GetValue(4);
                        var впростое = reader.GetValue(5);
                        var колвоЧасовВМесяце = reader.GetValue(6);
                        var процнаботыстанка = reader.GetValue(7);

                        dataSmpos.Add(new SMPO
                        {
                            date_start = date_start.ToString(),
                            IdStanok = Convert.ToInt32(IdStanok),
                            name_stan = name_stan,
                            ВНаладке = вналадке.ToString(),
                            ВПростое = впростое.ToString(),
                            ВРаботе = вработе.ToString(),
                            КолвоЧасовВМесяце = Convert.ToInt32(колвоЧасовВМесяце),
                            ПроцРаботыСтанка = процнаботыстанка.ToString()
                        });

                    }
                }
            }

            return dataSmpos;
        }
    }
}
