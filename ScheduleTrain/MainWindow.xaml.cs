using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScheduleTrain
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //PrintToXls();
        }

        public void PrintToXls()
        {
            var app = new Excel.Application();

            app.SheetsInNewWorkbook = 1;

            Excel.Workbook workbook = app.Workbooks.Add();
            Excel.Worksheet worksheet = app.Worksheets.Item[1];
            worksheet.Name = "Хуета";

            int startRow = 1;

            worksheet.Cells[1][startRow] = "хуй";
            worksheet.Cells[2][startRow] = "хуй";
            worksheet.Cells[3][startRow] = "хуй";
            worksheet.Cells[4][startRow] = "хуй";

            app.Visible = true;
        }

        public void PrintToCsv()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "info.csv";
            saveFileDialog.Filter = ".csv | *.csv ";

            // var biomaterials = Context._con.Biomaterial.ToList().Select(p => $"{p.Patient.InsuranceCompany.Name};{p.Patient.GetName};{p.GetServices};{p.GetPrice};{p.GetTotalPrice}").ToList();


        }
    }
}
