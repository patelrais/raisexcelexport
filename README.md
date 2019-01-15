 
        private void Form1_Load(object sender, EventArgs e)
        {

            List<Object> Employees = new List<object>();
            for (int i = 1; i <= 10; i++)
            {
                Employees.Add(new Employee { SrNo = i, Name = null, JoiningDate = DateTime.Now.AddDays(i), Hike = 12.34 + i, Score = (i) - 100 });
            }

             RaisExcelExport exportObject = new RaisExcelExport();

            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "SrNo", DataType = ExcelColumnDataType.Integer, HeaderText = "Sr. No.", Width = 20 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "Name", DataType = ExcelColumnDataType.Text, HeaderText = "Employee Name", Width = 50 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "JoiningDate", DataType = ExcelColumnDataType.Date, HeaderText = "Joining Date", Width = 15 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "Hike", DataType = ExcelColumnDataType.Double, HeaderText = "Hike", Width = 10 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "Score", DataType = ExcelColumnDataType.IntegerNegative, HeaderText = "Score", Width = 10 });

            exportObject.Data = Employees;
            exportObject.SheetName = "Employee List";

            /* Save to disk */
            exportObject.CreateExcelExport(@"c:\temp\EmployeeExport.xlsx");


            /* return as memorystream */
            List<Employee> data = new List<Employee>();
            MemoryStream memoryStream = exportObject.CreateReport();

            System.Diagnostics.Process.Start(@"c:\temp\EmployeeExport.xlsx");

        }
    }


    public class Employee
    {
        public int SrNo { get; set; }
        public string Name { get; set; }
        public DateTime JoiningDate { get; set; }
        public double Score { get; set; }
        public double Hike { get; set; }
    }
 
