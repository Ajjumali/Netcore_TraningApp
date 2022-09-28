using ClosedXML.Excel;
using DevExtreme.AspNet.Data;
using DevExtreme.AspNet.Mvc;
using DocumentFormat.OpenXml.Spreadsheet;
using Grpc.Core;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using TrainingApp.Models;


namespace TrainingApp.Controllers
{
    public class MileStoneController : Controller
    {
        MyConnection con = new MyConnection();
        [Obsolete]
        private IHostingEnvironment _hostingEnv;
        [Obsolete]
        public MileStoneController(IHostingEnvironment hostingEnv)
        {
            _hostingEnv = hostingEnv;
        }
        public object Get(DataSourceLoadOptions loadOptions)
        {
            return DataSourceLoader.Load(con.MileStone, loadOptions);
        }
        public object Fwbsdrop(DataSourceLoadOptions loadOptions)
        {

            return DataSourceLoader.Load(con.Fwbs, loadOptions);

        }
        public object pcwbsdrop(DataSourceLoadOptions loadOptions)
        {

            return DataSourceLoader.Load(con.PCWBS, loadOptions);
        }
        public ActionResult Index()
        {
            return View();
            
        }
        
        [HttpPost]
        public ActionResult AddMilestone(string values)
        {
            MyConnection con = new MyConnection();
            var newpcwbs = new MileStone();
            JsonConvert.PopulateObject(values, newpcwbs);
            if (!TryValidateModel(newpcwbs))
                return BadRequest(ModelState.ToString());

            con.MileStone.Add(newpcwbs);
            con.SaveChanges();
            return Ok();
        }
       
        [HttpPut]
        public ActionResult Editdata(string key, string values)
        {
            MyConnection con = new MyConnection();
            var employee = con.MileStone.FirstOrDefault(a => a.MileStone_Id == Convert.ToInt32(key));
            JsonConvert.PopulateObject(values, employee);

            if (!TryValidateModel(employee))
                return BadRequest(ModelState.ToString());

            con.SaveChanges();
            return Ok();
        }
        
        [HttpDelete]
        public void DeleteData(string key)
        {
            MyConnection con = new MyConnection();
            var employee = con.MileStone.FirstOrDefault(a => a.MileStone_Id == Convert.ToInt32(key));
            con.MileStone.Remove(employee);
            con.SaveChanges();
        }
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile postedFile)
        {
              //get file name
                var filename = ContentDispositionHeaderValue.Parse(postedFile.ContentDisposition).FileName.Trim('"');

                //get path
                var MainPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Uploads");

                //create directory "Uploads" if it doesn't exists
                if (!Directory.Exists(MainPath))
                {
                    Directory.CreateDirectory(MainPath);
                }

                //get file path 
                var filePath = Path.Combine(MainPath, postedFile.FileName);
                using (System.IO.Stream stream = new FileStream(filePath, FileMode.Create))
                {
                    await postedFile.CopyToAsync(stream);
                }

                //get extension
                string extension = Path.GetExtension(filename);


                string conString = string.Empty;

                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                }

                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }
                }
                //your database connection string
                conString = @"Server=172.16.41.226;Database=traningdb;User ID=CMS; password=jgc#2018$Dev;";

                using (SqlConnection con = new SqlConnection(conString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name.
                        sqlBulkCopy.DestinationTableName = "dbo.MileStone";

                        // Map the Excel columns with that of the database table, this is optional but good if you do
                        // 
                        sqlBulkCopy.ColumnMappings.Add("MileStone_Id", "MileStone_Id");
                        sqlBulkCopy.ColumnMappings.Add("ProjetName", "ProjetName");
                        sqlBulkCopy.ColumnMappings.Add("MileStone_Code", "MileStone_Code");
                        sqlBulkCopy.ColumnMappings.Add("MileStone_Descr", "MileStone_Descr");
                        sqlBulkCopy.ColumnMappings.Add("FKFwbs", "FKFwbs");
                        sqlBulkCopy.ColumnMappings.Add("FKPcwbs", "FKPcwbs");

                        con.Open();
                        sqlBulkCopy.WriteToServerAsync(dt);
                        con.Close();
                    }
                }
                //if the code reach here means everthing goes fine and excel data is imported into database
                //ViewBag.Message = "File Imported and excel data saved into database";
            

            return RedirectToAction("Index");

        }
        
        public ActionResult Download()
        {
            string Files = "wwwroot/Uploads/Testing.xlsx";
            byte[] fileBytes = System.IO.File.ReadAllBytes(Files);
            System.IO.File.WriteAllBytes(Files, fileBytes);
            MemoryStream ms = new MemoryStream(fileBytes);
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "MileStone.xlsx");
        }
        public IActionResult DownloadReport()
        {
            DataTable dt = new DataTable("MileStone");
            dt.Columns.AddRange(new DataColumn[6] { 
                                        new DataColumn("MileStone_Id"),
                                        new DataColumn("ProjetName"),
                                        new DataColumn("MileStone_Code"),
                                        new DataColumn("MileStone_Descr"),
                                        new DataColumn("FKFwbs"),
                                        new DataColumn("FKPcwbs"),
            });

            var customers = con.MileStone.ToList();
                            
                            
                            

            foreach (var customer in customers)
            {
                dt.Rows.Add(
                    customer.MileStone_Id,
                     customer.ProjetName,
                    customer.MileStone_Code,
                    customer.MileStone_Descr,
                    customer.FKFwbs,
                    customer.FKPcwbs);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MileStone.xlsx");
                }
            }
        }
    }
    

    
    
}



