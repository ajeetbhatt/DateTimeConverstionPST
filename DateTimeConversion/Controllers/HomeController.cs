using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using DateTimeConversion.Models;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;

namespace DateTimeConversion.Controllers
{
   
        public class HomeController : Controller
        {
            private const string SessionKey = "Tickets";
            public IActionResult Index()
            {
                return View();
            }
            /// <summary>
            /// This method is used to generate xls template
            /// </summary>
            /// <returns>it returns xls file</returns>
            [HttpGet]
            public ActionResult GenerateTemplate()
            {
                var dt = new DataTable("Grid");
                dt.Columns.AddRange(new[]
                {
                new DataColumn("TicketId*"),
                new DataColumn("CreatedDate"),
                new DataColumn("ModifiedDate"),
                new DataColumn("ClosedDate")

            });

                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add(dt);
                    var firstOrDefault = ws.Tables.FirstOrDefault();
                    if (firstOrDefault != null) firstOrDefault.ShowAutoFilter = false;
                    using (var stream = new MemoryStream())
                    {
                        Response.Headers.Add("content-disposition", "attachment;filename=Template_ticket_upload.xlsx");
                        wb.SaveAs(stream);
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "Template_ticket_upload.xlsx");
                    }
                }
            }
            /// <summary>
            /// This method is used to convert time to pst
            /// </summary>
            /// <param name="files"></param>
            /// <returns>it sets the converted ticket to session and return reponse in json format</returns>
            [HttpPost]
            public IActionResult UploadFile(List<IFormFile> files)
            {

                var dt = ImportExcelIntoDataTable(files);
                var tickets = new List<Ticket>();


                if (dt.Rows.Count > 0)
                {
                    for (var i = 0; i < dt.Rows.Count; i++)
                    {
                        Ticket ticket = new Ticket
                        {
                            TicketId = Convert.ToInt32(dt.Rows[i]["TicketId"].ToString()),
                            CreatedDateString = dt.Rows[i]["CreatedDate"].ToString(),
                            ModifiedDateString = dt.Rows[i]["ModifiedDate"].ToString(),
                            ClosedDateString = dt.Rows[i]["ClosedDate"].ToString()
                        };
                        tickets.Add(ticket);

                    }
                }

                if (tickets.Any())
                {
                    var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");

                    foreach (var ticket in tickets)
                    {

                        if (!string.IsNullOrEmpty(ticket.CreatedDateString) &&
                            DateTime.TryParse(ticket.CreatedDateString, out var createdDate))
                        {
                            ticket.CreatedDateTimeUtc =
                                TimeZoneInfo.ConvertTimeToUtc(createdDate, timeZone);
                        }
                        if (!string.IsNullOrEmpty(ticket.ModifiedDateString) &&
                            DateTime.TryParse(ticket.ModifiedDateString, out var modifiedDate))
                        {
                            ticket.ModifiedDateTimeUtc =
                                TimeZoneInfo.ConvertTimeToUtc(modifiedDate, timeZone);
                        }
                        if (!string.IsNullOrEmpty(ticket.ClosedDateString) &&
                            DateTime.TryParse(ticket.ClosedDateString, out var closeDate))
                        {
                            ticket.CloseDateTimeUtc =
                                TimeZoneInfo.ConvertTimeToUtc(closeDate, timeZone);
                        }

                    }
                    var serializeObject = JsonConvert.SerializeObject(tickets);
                    HttpContext.Session.SetString(SessionKey, serializeObject);

                }


                return Json(
                        new
                        {
                            success = true,
                            responseText = "Ticket time converted successfully."
                        });


            }

            [HttpGet]
            public ActionResult DownloadFile()
            {
                var value = HttpContext.Session.GetString(SessionKey);
                if (string.IsNullOrEmpty(value)) return Content("No Ticket FOund");
                List<Ticket> tickets = JsonConvert.DeserializeObject<List<Ticket>>(value);
                var returnDataTable = new DataTable("Grid");
                HttpContext.Session.Remove(SessionKey);
                returnDataTable.Columns.AddRange(new[]
                {
                new DataColumn("TicketId"),
                new DataColumn("CreatedDateUtc"),
                new DataColumn("CreatedDatePST"),
                new DataColumn("ModifiedDateUtc"),
                new DataColumn("ModifiedDatePST"),
                new DataColumn("CloseDateUtc"),
                new DataColumn("CloseDatePST"),
                new DataColumn("Mongo Script to Run")
            });
                foreach (var ticket in tickets)
                {
                    var createdDateStr = "";
                    var modifiedDateStr = "";
                    var closeDate = "";
                    var createdDate = "";
                    var modifiedDate = "";
                    var dbScript = "db.Tickets.updateOne({_id:" + ticket.TicketId + "},{$set:{";
                    if (ticket.CreatedDateTimeUtc > DateTime.MinValue && !string.IsNullOrEmpty(ticket.CreatedDateString))
                    {
                        createdDate = ticket.CreatedDateTimeUtc.ToString("s");
                        createdDateStr += "\"CreatedDateTimeUtc\":new ISODate(\"" + ticket.CreatedDateTimeUtc.ToString("s") + "\")";
                        dbScript += createdDateStr;
                    }
                    if (ticket.ModifiedDateTimeUtc > DateTime.MinValue && !string.IsNullOrEmpty(ticket.ModifiedDateString))
                    {
                        modifiedDate = ticket.ModifiedDateTimeUtc.ToString("s");
                        modifiedDateStr = "\"ModifiedByDateTimeUtc\":new ISODate(\"" + ticket.ModifiedDateTimeUtc.ToString("s") + "\")";
                        dbScript += !string.IsNullOrEmpty(createdDateStr) ? "," + modifiedDateStr : modifiedDateStr;
                    }
                    if (ticket.CloseDateTimeUtc > DateTime.MinValue && !string.IsNullOrEmpty(ticket.ClosedDateString))
                    {
                        var closeDateStr = "\"ClosedDateTimeUtc\":new ISODate(\"" + ticket.CloseDateTimeUtc.ToString("s") + "\")";
                        if (!string.IsNullOrEmpty(createdDateStr) || !string.IsNullOrEmpty(modifiedDateStr))
                        {
                            dbScript += "," + closeDateStr;
                        }
                        else
                        {
                            dbScript += closeDateStr;
                        }
                    }

             
                    dbScript += "}})";
              returnDataTable.Rows.Add(ticket.TicketId, createdDate, ticket.CreatedDateString, modifiedDate, ticket.ModifiedDateString,
                        closeDate, ticket.ClosedDateString, dbScript);
                }
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add(returnDataTable);
                    var firstOrDefault = ws.Tables.FirstOrDefault();
                    if (firstOrDefault != null) firstOrDefault.ShowAutoFilter = false;
                    using (var stream = new MemoryStream())
                    {
                        wb.SaveAs(stream);
                        var fileName = $"Tickets{DateTime.Now.Date.ToShortDateString().Replace("/", "-")}_.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            fileName);
                    }
                }
            }
            private DataTable ImportExcelIntoDataTable(IEnumerable<IFormFile> files)
            {
                var dt = new DataTable();
                var httpPostedFileBases = files as IFormFile[] ?? files.ToArray();
                if (!httpPostedFileBases.Any()) return dt;
                using (var workBook = new XLWorkbook(httpPostedFileBases.ToArray()[0].OpenReadStream()))
                {
                    var workSheet = workBook.Worksheet(1);
                    bool firstRow = true;
                    foreach (var row in workSheet.Rows())
                    {
                        if (row.IsEmpty()) continue;
                        if (firstRow)
                        {
                            foreach (var cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString().Replace("*", "").Trim());
                            }

                            firstRow = false;
                        }
                        else
                        {
                            dt.Rows.Add();
                            int i = 0;
                            foreach (var cell in row.Cells(workSheet.FirstCellUsed().Address.ColumnNumber,
                                workSheet.LastCellUsed().Address.ColumnNumber))
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }

                    }
                }

                return dt;
            }

        }
    }
