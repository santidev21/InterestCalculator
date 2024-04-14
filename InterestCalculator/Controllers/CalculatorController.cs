using InterestCalculator.Models;
using Microsoft.AspNetCore.Mvc;
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
    }
}
