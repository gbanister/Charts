﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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