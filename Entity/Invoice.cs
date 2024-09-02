using Microsoft.AspNetCore.Mvc;

namespace FileConversion.Entity
{
    public class Invoice
    {
        public int InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string CustomerName { get; set; }
        public string BillingAddress { get; set; }
        public string Description { get; set; }

        public double ContactNumber { get; set; }

        public string Email {  get; set; }
        public int Quantity { get; set; }
        public int UnitPrice{ get; set; }

        public double TotalPrice { get; set; }
        //public double TaxRate { get; set; }
        //public float TaxAmount{ get; set; }
        //public float TotalAmount { get; set; }

        public List<Product> Products { get; set; }


    }
}
