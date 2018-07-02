#region References
using ClosedXML.Excel;
using SampleApproval.Business.Members;
using SampleApproval.Model;
using SampleApproval.Web.ControllerAttributes;
using SampleApproval.Web.Filters;
using SampleApproval.Web.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Mvc;
#endregion

namespace SampleApproval.Web.Controllers
{
    [Authorize]
    [SessionManagement]
    [PasswordChange]
    [Administrator]
    public class ManageLabController : Controller
    {
        #region View Lab Info Actions
        /// <summary>
        /// GET: ManageLab
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public ActionResult Index(string message)
        {
            ViewBag.StatusMessage = message;

            return View();
        }
        #endregion

        #region Create New Lab Info Actions
        /// <summary>
        /// GET: ManageLab/Create
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        public ActionResult Create(string message)
        {
            ViewBag.StatusMessage = message;

            return View();
        }

        /// <summary>
        /// POST: ManageLab/Create
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(ManageLabModel model)
        {
            string error = string.Empty;

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            try
            {
                if (LabMember.Add(model.LabName, model.LabNameShort, model.Sequence, BaseContext.CurrentUser.UserName, out error))
                    return RedirectToAction("Index", new { Message = string.Format("Lab '{0}' successfully created.", model.LabName) });
                else
                    return RedirectToAction("Create", new { Message = error });
            }
            catch (Exception ex)
            {
                return RedirectToAction("Create", new { Message = ex.Message });
            }
        }
        #endregion

        #region Edit Existing Lab Info Actions
        /// <summary>
        /// GET: ManageLab/Edit/5
        /// </summary>
        /// <param name="id"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public ActionResult Edit(int id, string message)
        {
            ViewBag.StatusMessage = message;

            ManageLabModel model = LabMember.Get(id);
            return View(model);
        }

        /// <summary>
        /// POST: ManageLab/Edit/5
        /// </summary>
        /// <param name="id"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Edit(int id, ManageLabModel model)
        {
            string error = string.Empty;

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            try
            {
                if (LabMember.Edit(id, model.LabName, model.LabNameShort, model.Sequence, BaseContext.CurrentUser.UserName, out error))
                    return RedirectToAction("Index", new { Message = string.Format("Lab '{0}' successfully edited.", model.LabName) });
                else
                    return RedirectToAction("Edit", new { Message = error });
            }
            catch (Exception ex)
            {
                return RedirectToAction("Edit", new { Message = ex.Message });
            }
        }
        #endregion

        #region Delete Existing Lab Info Actions
        /// <summary>
        /// GET: ManageLab/Delete/5
        /// </summary>
        /// <param name="id"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public ActionResult Delete(int id, string message)
        {
            ViewBag.StatusMessage = message;

            ManageLabModel model = LabMember.Get(id);
            return View(model);
        }

        /// <summary>
        /// POST: ManageLab/Delete/5
        /// </summary>
        /// <param name="id"></param>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Delete(int id, ManageLabModel model)
        {
            string error = string.Empty;

            try
            {
                model = LabMember.Get(id);
                if (LabMember.Delete(id, BaseContext.CurrentUser.UserName, out error))
                    return RedirectToAction("Index", new { Message = string.Format("Lab '{0}' successfully deleted.", model.LabName) });
                else
                    return RedirectToAction("Delete", new { Message = error });
            }
            catch (Exception ex)
            {
                return RedirectToAction("Delete", new { Message = ex.Message });
            }
        }
        #endregion

        #region ExportToExcel Action
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ActionResult ExportData()
        {
            List<ManageLabModel> list = LabMember.Get();

            System.Data.DataTable tblData = new System.Data.DataTable();
            tblData.Columns.Add("Lab ID");
            tblData.Columns.Add("Lab Name");
            tblData.Columns.Add("Sequence");

            foreach (ManageLabModel item in list)
            {
                System.Data.DataRow row = tblData.NewRow();

                row[0] = item.LabID;
                row[1] = item.LabName;
                row[2] = item.Sequence;

                tblData.Rows.Add(row);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(tblData, "Lab");

                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Style.Font.Bold = true;

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=ExtractLabs.xlsx");

                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }

            return RedirectToAction("Index", "ManageLab");
        }
        #endregion

        #region Public Helper Methods
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetData()
        {
            return Json(LabMember.Get());
        }
        #endregion
    }
}