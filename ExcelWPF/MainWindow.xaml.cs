using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
using System.Windows.Xps.Packaging;

namespace ExcelWPF
{
 
    public partial class MainWindow : System.Windows.Window
    {

        string path = @"C:\Users\User\source\repos\ExcelWPF\doc.xlsx"; //поменяйте путь
        string xpsFileName = "";

        public MainWindow(){
            InitializeComponent();
            Loaded += MainWindow_Loaded;  
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e){
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(path);
            ExportXPS(excelWorkbook);
            excelWorkbook.Close(false, null, null);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
            excelApp = null;
            DisplayXPSFile();
        }

        
        void ExportXPS(Microsoft.Office.Interop.Excel.Workbook excelWorkbook){
            xpsFileName = (new DirectoryInfo(path)).FullName;
            xpsFileName = xpsFileName.Replace(new FileInfo(path).Extension, "") + ".xps";
            excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypeXPS,
            Filename: xpsFileName,
            OpenAfterPublish: false);
        }

        void DisplayXPSFile(){
            XpsDocument xpsPackage = new XpsDocument(xpsFileName, FileAccess.Read, CompressionOption.NotCompressed);
            FixedDocumentSequence fixedDocumentSequence = xpsPackage.GetFixedDocumentSequence();
            documentViewer.Document = fixedDocumentSequence;
        }


    }
}
