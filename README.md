# RaisExcelExport
  
Create excel export as memory stream or physical file.

# New Features!
  - Excel export in memorystream.
  - Save export to disk in .xlsx format.
  - HeaderRow background and foreground color.

```
using Rais;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        { 
            List<Object> Employees = new List<object>();
            for (int i = 1; i <= 10; i++)
            {
                Employees.Add(new Employee { SrNo = i, Name = null, JoiningDate = DateTime.Now.AddDays(i), Hike = 12.34 + i, Score = (i) - 100 });
            }

             RaisExcelExport exportObject = new RaisExcelExport();

            /* Use this link to convert RGB to Hex color : https://www.w3schools.com/colors/colors_rgb.asp */
            exportObject.DateFormat = "[$-409]d\\-mmm\\-yy;@";
            exportObject.HeaderBackgroundColorHex = "808080";
            exportObject.HeaderForeGround = "ffffff";
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "SrNo", DataType = ExcelColumnDataType.Integer, HeaderText = "Sr. No.", Width = 20 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "Name", DataType = ExcelColumnDataType.Text, HeaderText = "Employee Name", Width = 50 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "JoiningDate", DataType = ExcelColumnDataType.Date, HeaderText = "Joining Date", Width = 15 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "Hike", DataType = ExcelColumnDataType.Double, HeaderText = "Hike", Width = 10 });
            exportObject.Columns.Add(new ExcelColumn() { BindToProperty = "Score", DataType = ExcelColumnDataType.IntegerNegative, HeaderText = "Score", Width = 10 });

            exportObject.ExportList = Employees;
            exportObject.SheetName = "Employee List";

            /* Save to disk */
            exportObject.CreateExcelExport(@"c:\temp\EmployeeExport.xlsx");
              
            List<Employee> data = new List<Employee>();
            MemoryStream memoryStream = exportObject.CreateReportAsMemoryStream();

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
}


```


# License


MIT


**Free Library**

[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)


   [dill]: <https://github.com/joemccann/dillinger>
   [git-repo-url]: <https://github.com/joemccann/dillinger.git>
   [john gruber]: <http://daringfireball.net>
   [df1]: <http://daringfireball.net/projects/markdown/>
   [markdown-it]: <https://github.com/markdown-it/markdown-it>
   [Ace Editor]: <http://ace.ajax.org>
   [node.js]: <http://nodejs.org>
   [Twitter Bootstrap]: <http://twitter.github.com/bootstrap/>
   [jQuery]: <http://jquery.com>
   [@tjholowaychuk]: <http://twitter.com/tjholowaychuk>
   [express]: <http://expressjs.com>
   [AngularJS]: <http://angularjs.org>
   [Gulp]: <http://gulpjs.com>

   [PlDb]: <https://github.com/joemccann/dillinger/tree/master/plugins/dropbox/README.md>
   [PlGh]: <https://github.com/joemccann/dillinger/tree/master/plugins/github/README.md>
   [PlGd]: <https://github.com/joemccann/dillinger/tree/master/plugins/googledrive/README.md>
   [PlOd]: <https://github.com/joemccann/dillinger/tree/master/plugins/onedrive/README.md>
   [PlMe]: <https://github.com/joemccann/dillinger/tree/master/plugins/medium/README.md>
   [PlGa]: <https://github.com/RahulHP/dillinger/blob/master/plugins/googleanalytics/README.md>
