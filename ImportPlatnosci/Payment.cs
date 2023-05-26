using Soneta.Types;

namespace ImportPlatnosci
{
    public class Payment
    {
        public Date Date { get; }
        public string Id { get; }
        public Currency Amount { get; }
        public string Description { get; set; }
        public string PaymentType { get; }
        public string Contractor { get; }

        public Payment(Date date, string id, Currency amount, string desc, string paymentType)
        {
            Date = date;
            Id = id;
            Amount = amount;
            Description = desc;
            PaymentType = paymentType;
        }

        public Payment(Date date, string id, Currency amount, string desc, string paymentType, string contractor)
        {
            Date = date;
            Id = id;
            Amount = amount;
            Description = desc;
            PaymentType = paymentType;
            Contractor = contractor;
        }
    }
}
