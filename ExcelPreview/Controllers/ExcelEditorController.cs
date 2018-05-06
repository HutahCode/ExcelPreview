using ExcelPreview.DataRepo;
using ExcelPreview.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelPreview.Controllers
{
    public class ExcelEditorController : Controller
    {
        ExcelEditorRepo _excelEditorRepo = null;
        public ExcelEditorController()
        {
            _excelEditorRepo = new ExcelEditorRepo();
        }

        // GET: ExcelEditor
        public ActionResult Index(string excelFile)
        {
            string extension = Path.GetExtension(excelFile);
            string excelConnectionString = string.Empty;
            switch (extension)
            {
                case ".xls":
                    excelConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = {0}; Extended Properties = 'Excel 8.0;HDR=YES'";
                    break;
                case ".xlsx":
                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                    break;
            }

            DataTable dtExcel = new DataTable();
            excelConnectionString = string.Format(excelConnectionString, excelFile);

            using (OleDbConnection connExcel = new OleDbConnection(excelConnectionString))
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
                        odaExcel.Fill(dtExcel);
                        connExcel.Close();
                    }
                }
            }

            var canidateList = Helper.DataTableMapping.CreateListFromTableForAll<CandidateInfo>(dtExcel);
                        
            return View(canidateList);
        }

       

        [HttpPost]
        public ActionResult Index(IEnumerable<CandidateInfo> candidateList)
        {
            if (ModelState.IsValid)
            {
                var dtExcel = Helper.DataTableMapping.CreateDataTableFromList<CandidateInfo>(candidateList.ToList());
                string excelFilePath = Server.MapPath("~/DownloadExcel/CandidateInfo.xlsx");
                _excelEditorRepo.CreateExcelDoc(excelFilePath, candidateList.ToList());

                byte[] fileBytes = System.IO.File.ReadAllBytes(excelFilePath);
                string fileName = Path.GetFileName(excelFilePath);
                return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
            }
            
            return View(candidateList);
        }
        

    }
}