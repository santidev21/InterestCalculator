namespace InterestCalculator.Models
{
    public class Payment
    {
        public int InstallmentNumber { get; set; }
        public decimal PrincipalPaid { get; set; }
        public decimal InterestPaid { get; set; }
        public decimal RemainingPrincipal { get; set; }
        public decimal PaymentAmount { get; set; }
        public DateTime DueDate { get; set; }
    }
}
