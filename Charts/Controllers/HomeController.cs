using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Helpers;
using System.Web.Mvc;
using Charts.Models;
using Excel;
using Newtonsoft.Json;

namespace Charts.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }


        [HttpGet]
        public JsonResult GetLineChartData()
        {
            var filePath2010 = HttpContext.Server.MapPath("~/App_Data/Announced Deals w # 2010 v02.xls");
            var filePath2011 = HttpContext.Server.MapPath("~/App_Data/Announced Deals w # 2011 v02.xls");
            var table = InitTable();
            GetQarterlyDeals(filePath2010, table);
            GetQarterlyDeals(filePath2011, table);
            CalculateMeans(table);
            var json = JsonConvert.SerializeObject(table, Formatting.None);
            return Json(json, JsonRequestBehavior.AllowGet);
        }

        private void CalculateMeans(object[,] table)
        {
            for (var row = 1; row < 9; row++)
                for (var col = 1; col < 5; col++)
                {
                    var items = (List<int>) table[row, col];
                    if (items.Count == 0)
                    {
                        table[row, col] = 0;
                    }
                    else
                    {
                        table[row, col] = items.Average();
                    }
                }
        }

        private void GetQarterlyDeals(string filePath, object[,] table)
        {
            FileStream stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;


            var isFirstRow = true;
            while (excelReader.Read())
            {
                if (isFirstRow)
                {
                    isFirstRow = false;
                    continue;
                }

                var value = excelReader.GetDecimal(7);
                if (value < 100) continue;

                var numAdvisors = excelReader.GetInt32(12);
                if (numAdvisors < 1) continue;
                
                var date = excelReader.GetDateTime(0);
                var quarter = GetQuarter(date);
                var tableRow = (date.Year - 2009)*quarter;
                var tableCol = GetValueColumn(value);
                ((List<int>) table[tableRow, tableCol]).Add(numAdvisors);
            }

            excelReader.Close();
            excelReader.Dispose();
        }

        private object[,] InitTable()
        {
            var table = new object[9, 5];
            table[0, 0] = "Quarter";
            table[0, 1] = "100(mil)";
            table[0, 2] = "500(mil)";
            table[0, 3] = "1000(mil)";
            table[0, 4] = "5000(mil)";
            table[1, 0] = "2010 Q1";
            table[2, 0] = "2010 Q2";
            table[3, 0] = "2010 Q3";
            table[4, 0] = "2010 Q4";
            table[5, 0] = "2011 Q1";
            table[6, 0] = "2011 Q2";
            table[7, 0] = "2011 Q3";
            table[8, 0] = "2011 Q4";

            for (var row = 1; row < 9; row++)
                for (var col = 1; col < 5; col++)
                {
                    table[row, col] = new List<int>();
                }
            return table;
        }

        private int GetValueColumn(decimal value)
        {
            if (value < 500) return 1;
            if (value < 1000) return 2;
            if (value < 5000) return 3;
            return 4;
        }

        public static int GetQuarter(DateTime date)
        {
            if (date.Month >= 4 && date.Month <= 6)
                return 1;
            else if (date.Month >= 7 && date.Month <= 9)
                return 2;
            else if (date.Month >= 10 && date.Month <= 12)
                return 3;
            else
                return 4;

        }

        [HttpGet]
        public JsonResult GetPieChartData(string advisors)
        {
            var filePath = HttpContext.Server.MapPath("~/App_Data/Announced Deals w # 2011 v02.xls");

            string[] advisorArray;
            try
            {
                advisorArray = advisors.Split(',');
            }
            catch (Exception)
            {
                return Json(new {Error ="bad input" }, JsonRequestBehavior.AllowGet);
            }

            var advisorDeals  = GetIndustryByAdvisor(filePath, advisorArray);
  
            object[,] table = new object[advisorDeals.Count+1, 2];
            table[0, 0] = "Industry";
            table[0, 1] = "Deals";

            for (int i = 0; i < advisorDeals.Count; i++)
            {
                table[i+1, 0] = advisorDeals[i].Industry;
                table[i+1, 1] = advisorDeals[i].Deals;
            }

            var json = JsonConvert.SerializeObject(table, Formatting.None);

            return Json(json, JsonRequestBehavior.AllowGet);
        }



        [HttpGet]
        public JsonResult GetChartData()
        {
            var filePath = HttpContext.Server.MapPath("~/App_Data/Announced Deals w # 2010 v02.xls");
            // var filePath = @"\App_Data\Announced Deals w # 2010 v02.xls";

            var sums1 = GetRegionDeals(filePath);
            filePath = HttpContext.Server.MapPath("~/App_Data/Announced Deals w # 2011 v02.xls");
            var sums2 = GetRegionDeals(filePath);

            object[,] table = new object[3, 5];
            table[0, 0] = "API Category";
            table[0, 1] = sums1[0].Region;
            table[0, 2] = sums1[1].Region;
            table[0, 3] = sums1[2].Region;
            table[0, 4] = new { role = "annotation" };
            table[1, 0] = "2010";
            table[1, 1] = sums1[0].Value;
            table[1, 2] = sums1[1].Value;
            table[1, 3] = sums1[2].Value;
            table[1, 4] = "";
            table[2, 0] = "2011";
            table[2, 1] = sums2[0].Value;
            table[2, 2] = sums2[1].Value;
            table[2, 3] = sums2[2].Value;
            table[2, 4] = "";

            var json = JsonConvert.SerializeObject(table, Formatting.None);

            return Json(json, JsonRequestBehavior.AllowGet);
        }

        private static List<AdvisorDeal> GetIndustryByAdvisor(string filePath, params string[] advisorVariations)
        {
            FileStream stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            excelReader.IsFirstRowAsColumnNames = true;

            //5. Data Reader methods
            var deals = new Dictionary<string, int>();
           
            var isFirstRow = true;
            while (excelReader.Read())
            {
                if (isFirstRow)
                {
                    isFirstRow = false;
                    continue;
                }
                var acquirorAdvisor = excelReader.GetString(3) ?? "";
                var targetAdvisor = excelReader.GetString(4) ?? "";
                foreach (var advisor in advisorVariations)
                {
                    if (acquirorAdvisor.Contains(advisor) || targetAdvisor.Contains(advisor))
                    {
                        var industry = excelReader.GetString(8);
                        if (deals.ContainsKey(industry))
                        {
                            deals[industry] ++;
                        }
                        else
                        {
                            deals.Add(industry, 1);
                        }
                    }
                }

            }

            excelReader.Close();
            excelReader.Dispose();

            var advisorDeals = deals.Select(x => new AdvisorDeal { Deals = x.Value, Industry = x.Key })
                .OrderBy(x => x.Industry)
                .ToList();

            return advisorDeals;
        }


        private static List<RegionDeal> GetRegionDeals(string filePath)
        {
            FileStream stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);

            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            //...
            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();
            //...
            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            //           var dataSet = excelReader.AsDataSet();

            //5. Data Reader methods
            var deals = new List<RegionDeal>();
            var isFirstRow = true;
            while (excelReader.Read())
            {
                if (isFirstRow)
                {
                    isFirstRow = false;
                    continue;
                }
                var value = excelReader.GetDecimal(7);
                if (value < 0) value = 0; 
                deals.Add(new RegionDeal { Region = excelReader.GetString(6), Value = value });
            }

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
            excelReader.Dispose();


            //            var deals = new List<RegionDeal>();
            //            foreach (DataRow row in dataSet.Tables[0].Rows)
            //            {
            //                var value = row[7] is DBNull ? 0 : Convert.ToDecimal(row[7]);
            //                deals.Add(new RegionDeal {Region = row[6].ToString(), Value = value});
            //            }

            var sums = deals.GroupBy(d => d.Region)
                .Select(x => new RegionDeal { Value = x.Sum(d => d.Value), Region = x.First().Region})
                .OrderBy(x => x.Region)
                .ToList();

            return sums;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}