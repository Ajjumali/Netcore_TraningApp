using ClosedXML.Excel;
using DevExtreme.AspNet.Data;
using DevExtreme.AspNet.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TrainingApp.Models;

namespace TrainingApp.Controllers
{
    public class TestController : Controller
    {
        MyConnection con = new MyConnection();
        public object Get(DataSourceLoadOptions loadOptions)
        {
            return DataSourceLoader.Load(con.PCWBS, loadOptions);
        }
        public object drop(DataSourceLoadOptions loadOptions)
        {
            
            return DataSourceLoader.Load(con.PCWBS, loadOptions);
        }
        public ActionResult Index()
        {
            
            return View();
        }
        
        [HttpPost]
        public ActionResult AddPCwbc(string values)
        {
            MyConnection con = new MyConnection();
            var newpcwbs = new PCWPS();
            JsonConvert.PopulateObject(values,newpcwbs);
            if (!TryValidateModel(newpcwbs))
                return BadRequest(ModelState.ToString());

            con.PCWBS.Add(newpcwbs);
            con.SaveChanges();
            return Ok();
        }
        
        [HttpPut]
        public ActionResult Editdata(string key, string values)
        {
            MyConnection con = new MyConnection();
            var employee = con.PCWBS.FirstOrDefault(a => a.Pcwbs_Id == Convert.ToInt32(key));
            JsonConvert.PopulateObject(values, employee);

            if (!TryValidateModel(employee))
                return BadRequest(ModelState.ToString());

            con.SaveChanges();
            return Ok();

        }
        
        public void DeleteData(string key)
        {
            MyConnection con = new MyConnection();
             var employee = con.PCWBS.FirstOrDefault(a => a.Pcwbs_Id== Convert.ToInt32(key));
             con.PCWBS.Remove(employee);
            con.SaveChanges();

        }
        public IActionResult DownloadReport()
        {
            DataTable dt = new DataTable("PCWBS");
            dt.Columns.AddRange(new DataColumn[14] {
                                        new DataColumn("Pcwbs_Id"),
                                        new DataColumn("Job_Code_Id"),
                                        new DataColumn("Prime_Code"),
                                        new DataColumn("Pcwbs"),
                                        new DataColumn("Parent_Pcwbs_Id"),
                                        new DataColumn("Pcwbs_Descr"),
                                        new DataColumn("Pcwbs_Level_No"),
                                        new DataColumn("Priority_Num"),
                                        new DataColumn("Plan_Shop_Subcon"),
                                        new DataColumn("Plan_Field_Subcon"),
                                        new DataColumn("Active_Flg"),
                                        new DataColumn("Remarks"),
                                        new DataColumn("Modified_On"),
                                        new DataColumn("Created_On"),
            });

            var customers = con.PCWBS.ToList();

            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.Pcwbs_Id,
                    customer.Job_Code_Id,
                    customer.Prime_Code,
                    customer.Pcwbs,
                    customer.Parent_Pcwbs_Id,
                    customer.Pcwbs_Descr,
                    customer.Pcwbs_Level_No,
                    customer.Priority_Num,
                    customer.Plan_Shop_Subcon,
                    customer.Plan_Field_Subcon,
                    customer.Active_Flg,
                    customer.Remarks,
                    customer.Modified_On,
                    customer.Created_On
                    );
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PCWBS.xlsx");
                }
            }
        }
    }
}
