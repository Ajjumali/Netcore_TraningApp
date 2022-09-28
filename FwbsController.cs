using ClosedXML.Excel;
using DevExtreme.AspNet.Data;
using DevExtreme.AspNet.Mvc;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
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
    public class FwbsController : Controller
    {
        MyConnection con = new MyConnection();

        public object Get(DataSourceLoadOptions loadOptions)
        {
            
            return DataSourceLoader.Load(con.Fwbs, loadOptions);
        }
        public object drop(DataSourceLoadOptions loadOptions)
        {

            return DataSourceLoader.Load(con.Fwbs, loadOptions);
        }
        public ActionResult Index()
        {
           
            return View();
        }
        public ActionResult AddFwbs(string values)
        {
            var newpcwbs = new FWBS();
            JsonConvert.PopulateObject(values, newpcwbs);
            if (!TryValidateModel(newpcwbs))
                return BadRequest(ModelState.ToString());

            con.Fwbs.Add(newpcwbs);
            con.SaveChanges();
            return Ok();
        }
        [HttpPut]
        public ActionResult Editdata(string key, string values)
        {
            MyConnection con = new MyConnection();
            var employee = con.Fwbs.FirstOrDefault(a => a.Fwbs_Id == Convert.ToInt32(key));
            JsonConvert.PopulateObject(values, employee);

            if (!TryValidateModel(employee))
                return BadRequest(ModelState.ToString());

            con.SaveChanges();
            return Ok();

        }
        [HttpDelete]
        public void DeleteData(string key)
        {
            var employee = con.Fwbs.FirstOrDefault(a => a.Fwbs_Id == Convert.ToInt32(key));
            con.Fwbs.Remove(employee);
            con.SaveChanges();
        }
        public IActionResult DownloadReport()
        {
            DataTable dt = new DataTable("FWBS");
            dt.Columns.AddRange(new DataColumn[8] {
                                        new DataColumn("Fwbs_Id"),
                                        new DataColumn("ProjetName"),
                                        new DataColumn("Prime_Code"),
                                        new DataColumn("Fwbs"),
                                        new DataColumn("Parent_Fwbs_Id"),
                                        new DataColumn("Fwbs_Descr"),
                                        new DataColumn("Fwbs_Level_No"),
                                        new DataColumn("BQ_Unit"),
            });

            var customers = con.Fwbs.ToList();
            foreach (var customer in customers)
            {
                dt.Rows.Add(customer.Fwbs_Id,
                    customer.ProjectName,
                    customer.Prime_Code,
                    customer.Fwbs,
                    customer.Parent_Fwbs_Id,
                    customer.Fwbs_Descr,
                    customer.Fwbs_Level_No,
                    customer.BQ_Unit);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FWBS.xlsx");
                }
            }
        }
    }
}
