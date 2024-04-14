namespace InterestCalculator.Models
{
    public class CalculatorModel
    {
        public decimal PurchaseAmount { get; set; }
        public int Installments { get; set; }

        public decimal Interest { get; set; }
        public InterestType InterestType { get; set; }
        public List<Payment> PaymentPlan { get; set; } = new List<Payment>();


        public void CalculatePaymentPlan()
        {
            decimal monthlyRate;
            if( InterestType == InterestType.effectiveAnnualRate)
            {
                monthlyRate = (decimal)Math.Pow((double)(1 + (Interest/100)), 1.0 / 12) - 1;
            }
            else
            {
                monthlyRate = Interest / 100;
            }

            decimal monthlyPrincipalPayment = PurchaseAmount / Installments;

            decimal remainingAmount = PurchaseAmount;

            for (int i = 1; i <= Installments; i++)
            {
                decimal interestPaid = remainingAmount * monthlyRate;
                decimal paymentAmount = monthlyPrincipalPayment + interestPaid;
                remainingAmount = remainingAmount - monthlyPrincipalPayment;

                PaymentPlan.Add(new Payment
                {
                    InstallmentNumber = i,
                    PrincipalPaid = monthlyPrincipalPayment,
                    InterestPaid = interestPaid,
                    PaymentAmount = paymentAmount,
                    RemainingPrincipal = remainingAmount

                });
            }

        }



    }
}
