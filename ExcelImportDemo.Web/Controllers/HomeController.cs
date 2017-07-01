using ExcelImportDemo.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelImportDemo.Web.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase excelFile, string hdnTest)
        {
            var data = (new ExcelReader()).ReadExcel(excelFile);
            //Session[SessionExcelData] = data;

            return View("Index", data.Status);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            //    [HttpPost]
            //[ActionName(ActionNameConstants.ImportProject)]
            //public IHttpActionResult ImportProject(ImportProjectModel importProjectModel)
            //{
            //    var importStartDate = DateTime.UtcNow;
            //    var isAccepted = importProjectModel.FileName.Contains(ApplicationEnums.Accepted);

            //    try
            //    {
            //        int columnCount;

            //        var getTempRecord = Uow.Repository<ImportProjectScheduler_Temp>().All().ToList();
            //        if (getTempRecord.Count() > 0)
            //            Uow.Repository<ImportProjectScheduler_Temp>().DeleteRange(getTempRecord);


            //        var dt = ExcelHelperDomain.READExcel(importProjectModel.FileSavePath, out columnCount);
            //        for (var i = 0; i < dt.Rows.Count; i++)
            //        {
            //            var rowData = dt.Rows[i];
            //            DateTime? date = null;
            //            var orderAcceptanceDate = string.Empty;
            //            DateTime projectDate;
            //            DateTime orderDate = DateTime.UtcNow;
            //            var acceptanceDate = false;
            //            //var getOrderDate = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProjectCreateDate].ToString();
            //            //var getProjectCreatedDate = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProjectCreateDate].ToString();
            //            //if (getProjectCreatedDate == "")
            //            //    continue;
            //            //var projectDate = DateTime.FromOADate(Convert.ToDouble(getProjectCreatedDate));


            //            //var projectDate = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.OrderAcceptedDate].ToString();
            //            var projectCreatedDate = DateTime.TryParseExact(rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProjectCreateDate].ToString(), ApplicationConstants.DateFormat.YYYY_MM_dd, System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out projectDate);
            //            if (columnCount == ApplicationEnums.AcceptedImportColumns)
            //            {
            //                acceptanceDate = DateTime.TryParseExact(rowData.ItemArray[ApplicationEnums.ProjectImportColumns.OrderAcceptedDate].ToString(), ApplicationConstants.DateFormat.YYYY_MM_dd, System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out orderDate);
            //            }


            //            if ((acceptanceDate == false && isAccepted == true) || projectCreatedDate == false)
            //                return Ok<bool>(false);

            //            var scheduler_Temp = new ImportProjectScheduler_Temp
            //            {
            //                ProjectCreateDate = projectDate,
            //                ProjectNumber = Convert.ToInt32(rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProjectNumber]),
            //                ProjectName = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProjectName].ToString(),
            //                CreatorFirstName = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.CreatorFirstName].ToString(),
            //                CreatorLastName = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.CreatorLastName].ToString(),
            //                CreatorEmail = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.CreatorEmail].ToString(),
            //                ClientWorkgroupName = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ClientWorkgroupName].ToString(),
            //                SpecID = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.SpecID].ToString(),
            //                SpecName = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.SpecName].ToString(),
            //                Brand = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.Brand].ToString(),
            //                ProductType = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProductType].ToString(),
            //                UnitClass = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.UnitClass].ToString(),
            //                OrderQuantity = Convert.ToInt32(rowData.ItemArray[ApplicationEnums.ProjectImportColumns.OrderQuantity]),
            //                OrderStatus = rowData.ItemArray[ApplicationEnums.ProjectImportColumns.OrderStatus].ToString(),
            //                OrderID = isAccepted == true ? rowData.ItemArray[ApplicationEnums.ProjectImportColumns.OrderID].ToString() : "",
            //                //OrderAcceptedDate = getOrderDate != "" ? DateTime.FromOADate(Convert.ToDouble(rowData.ItemArray[ApplicationEnums.ProjectImportColumns.ProjectCreateDate])) : date,
            //                OrderAcceptedDate = acceptanceDate == true ? orderDate : date,
            //                OrderTransactionType = isAccepted == true ? rowData.ItemArray[ApplicationEnums.ProjectImportColumns.OrderTransactionType].ToString() : "",
            //                SupplierWorkgroupName = isAccepted == true ? rowData.ItemArray[ApplicationEnums.ProjectImportColumns.SupplierWorkgroupName].ToString() : "",
            //                //CreatedDate = Convert.ToDateTime(rowData.ItemArray[ApplicationEnums.ProjectImportColumns.CreatedDate])
            //                CreatedDate = importStartDate
            //            };
            //            Uow.Repository<ImportProjectScheduler_Temp>().Add(scheduler_Temp);
            //        }
            //        Uow.Save();

            //        if ((System.IO.File.Exists(importProjectModel.FileSavePath)))
            //            System.IO.File.Delete(importProjectModel.FileSavePath);

            //        var result = GetDataFromTemp(importProjectModel.FileName);

            //        if (result == "True")
            //            return Ok<bool>(true);
            //        else
            //            return Ok(result);
            //    }
            //    catch (Exception ex)
            //    {
            //        DatabaseLog(ex, importStartDate);
            //        return Ok<bool>(false);
            //    }
            //}

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}