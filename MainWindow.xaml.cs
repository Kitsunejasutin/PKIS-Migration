using Dapper;
using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Data.SqlClient;
using System.Windows;

namespace PKIS_Migration
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ReadExcel();
        }

        public void ReadExcel()
        {
            try
            {
                List<DataStructure> SKUs = [];
                List<DataStructure> UniqueSKUs = [];
                List<InvoiceStructure> Invoices = [];
                Dictionary<string, int> Company = new()
                {
                    ["companyName_1"] = 1,
                    ["companyName_2"] = 2,
                };

                using ExcelDataReader skus = ExcelDataReader.Create("path_to_excel.xlsx");
                foreach (DataStructure item in skus.GetRecords<DataStructure>())
                {
                    if (!UniqueSKUs.Any(sku => sku.item_sku == item.item_sku))
                        UniqueSKUs.Add(item);
                    SKUs.Add(item);
                }

                var connection = new SqlConnection("Server=.\\SQLEXPRESS;Database=pkisdb;User ID=pkis_user;Password=Pk1$+USER!");

                foreach (DataStructure item in UniqueSKUs)
                    connection.Execute("INSERT INTO dbo.item_sku (skuid, skudesc, skuagesex, skuitemsold, skuunitprice, skustyle, skudate) VALUES " +
                                       "(@item_sku, @item_description, @item_age_sex, @item_sold, @unit_price, @invoice_styles, @Now);",
                                       new { item.item_sku, item.item_description, item.item_age_sex, item.item_sold, item.unit_price, item.invoice_styles, DateTime.Now });

                using ExcelDataReader invoices = ExcelDataReader.Create("path_to_excel.xlsx");
                foreach (InvoiceStructure item in invoices.GetRecords<InvoiceStructure>())
                    if (!Invoices.Any(invoice => invoice.invoice_number == item.invoice_number))
                        Invoices.Add(item);

                foreach (InvoiceStructure item in Invoices)
                {
                    Company.TryGetValue(item.company_name, out int CompanyID);

                    connection.Execute("INSERT INTO dbo.invoices (inum, idate, cid, po_num) VALUES (@invoice_number, @invoice_date, @CompanyID, @purchase_order_number);",
                                       new { item.invoice_number, item.invoice_date, CompanyID, item.purchase_order_number });

                    foreach (DataStructure sku in SKUs.Where(sku => sku.invoice_number == item.invoice_number))
                    {
                        connection.Execute("INSERT INTO dbo.invoice_items (inum, skuid, iquantity) VALUES (@invoice_number, @item_sku, @item_quantity);",
                                           new { sku.invoice_number, sku.item_sku, sku.item_quantity });
                    }

                }
                MessageBox.Show("All Done!");
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        class DataStructure
        {
            public int invoice_number { get; set; }
            public int item_sku { get; set; }
            public string item_description { get; set; }
            public string item_age_sex { get; set; }
            public string item_sold { get; set; }
            public int item_quantity { get; set; }
            public int unit_price { get; set; }
            public int item_amount { get; set; }
            public string invoice_styles { get; set; }

        }
        class InvoiceStructure
        {
            public int invoice_number { get; set; }
            public DateTime invoice_date { get; set; }
            public string company_name { get; set; }
            public int purchase_order_number { get; set; }
        }
    }
}
