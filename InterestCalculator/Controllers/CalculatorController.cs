using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using InterestCalculator.Models;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Globalization;
using System.Reflection;

namespace InterestCalculator.Controllers
{
    public class CalculatorController : Controller
    {
        public IActionResult Index()
        {
            var model = new CalculatorModel();
            return View(model);
        }

        public IActionResult Calculate(CalculatorModel model) {

            // Fix decimal issue
            decimal interest;
            bool isSuccess = decimal.TryParse(model.Interest.ToString(), out interest);
            if (isSuccess)
            {
                model.Interest = interest;

                if (ModelState.IsValid)
                {
                    model.CalculatePaymentPlan();
                    return View("PaymentPlan", model);
                }
            }           

            return View("Index", model);
        }

        public IActionResult Export(CalculatorModel model)
        {
            decimal interest;
            bool isSuccess = decimal.TryParse(model.Interest.ToString(), out interest);
            if (isSuccess)
            {
                model.Interest = interest;
                model.CalculatePaymentPlan();
            }
            
            var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Payment plan");

            // Setup columns
            ws.Cell(1, 1).Value = "Installment";
            ws.Cell(1, 2).Value = "Principal paid";
            ws.Cell(1, 3).Value = "Interest paid";
            ws.Cell(1, 4).Value = "Payment amount";
            ws.Cell(1, 5).Value = "Remaining amount";

            setHeadersFormat(wb, ws);

            int currentRow = 2;
            // Filing rows
            foreach (Payment row in model.PaymentPlan)
            { 
                ws.Cell(currentRow, 1).Value = row.InstallmentNumber;
                ws.Cell(currentRow, 2).Value = row.PrincipalPaid;
                ws.Cell(currentRow, 3).Value = row.InterestPaid;
                ws.Cell(currentRow, 4).Value = row.PaymentAmount;
                ws.Cell(currentRow, 5).Value = row.RemainingPrincipal;
                currentRow++;
            }

            int finalRow = currentRow - 1;
            setDataFormat(wb, ws, finalRow.ToString());
            setTableBorderFormat(wb, ws, finalRow.ToString());

            var stream = new MemoryStream();
            wb.SaveAs(stream);
            stream.Position = 0;

            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "payment_plan.xlsx");
        }

        private void setHeadersFormat(IXLWorkbook wb, IXLWorksheet ws)
        {
            var range = ws.Range("A1:E1");

            range.Style.Font.Bold = true;

            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Columns().AdjustToContents();
        }

        private void setDataFormat(IXLWorkbook wb, IXLWorksheet ws, string finalRow)
        {
            var dataRange = ws.Range("A2:E" + finalRow);

            dataRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            dataRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            var currencyRange = ws.Range("B2:E" + finalRow);
            currencyRange.Style.NumberFormat.Format = "$#,##0;-$#,##0";

            ws.Columns().AdjustToContents();
        }

        private void setTableBorderFormat(IXLWorkbook wb, IXLWorksheet ws, string finalRow)
        {
            var range = ws.Range("A1:E"+finalRow);

            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.OutsideBorderColor = XLColor.Black;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorderColor = XLColor.Black;

            ws.Columns().AdjustToContents();
        }
    }
}
