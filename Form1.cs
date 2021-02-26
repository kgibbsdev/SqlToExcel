using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace ExcelPractice
{
    public partial class Form1 : Form
    {
        public List<string> products = new List<string>();
        public Form1()
        {
            InitializeComponent();
    }

        private void Form1_Load(object sender, EventArgs e)
        {
            //var test = new Product()
            //{
            //    ProductID = 1111,
            //    Name = "Rocket Fuel",
            //    ProductNumber = "RF-3227",
            //    MakeFlag = true,
            //    FinishedGoodsFlag = true,
            //    SafetyStockLevel = 10,
            //    ReorderPoint = 4,
            //    StandardCost = (decimal)8898.99,
            //    ListPrice = (decimal)25000.99,
            //    SellStartDate = DateTime.Now,
            //    ModifiedDate = DateTime.Now
            //};

            //context.Products.Add(test);
            //context.SaveChanges();
            button1.Enabled = false;
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Excel.Application excelApp;
            Excel._Workbook excelWorkBook;
            Excel._Worksheet excelSheet;

            try
            {
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                //Start Excel and get Application object.
                excelApp = new Excel.Application();
                
                //Get a new workbook.
                excelWorkBook = (Excel._Workbook)(excelApp.Workbooks.Add(Missing.Value));
                excelSheet = (Excel._Worksheet)excelWorkBook.ActiveSheet;

                //Add table headers going cell by cell.
                excelSheet.Cells[1, 1] = "Product Name";

                string cellName;
                int counter = 2;

                foreach (string productName in products)
                {
                    cellName = "A" + counter.ToString();
                    Excel.Range range = excelSheet.get_Range(cellName, cellName);
                    range.Value2 = productName;
                    counter++;
                }

                //Best practice to make the Excel sheet visible after the entire spreadsheet is ready.
                excelApp.Visible = true;

                stopWatch.Stop();
                Console.WriteLine("Time: " + stopWatch.ElapsedMilliseconds.ToString() + "ms");
            }
            catch(Exception exception)
            {
                Console.WriteLine("Exception Caught: "+ exception.Message + " " + exception.InnerException);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
             var context = new AdventureWorks2019Entities();
            
            foreach(Product product in context.Products)
            {
                products.Add(product.Name);
            }

            //foreach(string productName in products)
            //{
            //    Console.WriteLine(productName);
            //}

            button1.Enabled = true;
        }
    }
}
