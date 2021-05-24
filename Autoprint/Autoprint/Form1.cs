using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using RawPrint;

namespace Autoprint
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            richEditControl1.LoadDocument(@"C:\Users\DELL\Documents\Zalo Received Files\PO-210300062.pdf");
            gridControl1.ShowPrintPreview();
           //richEditControl1.Print();
        }

        private void print()
        {
            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Load Excel workbook from file's path.
            ExcelFile workbook = ExcelFile.Load(@"E:\TECKSOL\Tailieu\BAOCAO.xlsx");

            // Set sheets print options.
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                ExcelPrintOptions sheetPrintOptions = worksheet.PrintOptions;

                sheetPrintOptions.Portrait = false;
                sheetPrintOptions.HorizontalCentered = true;
                sheetPrintOptions.VerticalCentered = true;

                sheetPrintOptions.PrintHeadings = true;
                sheetPrintOptions.PrintGridlines = true;
            }

            // Create spreadsheet's print options. 
            PrintOptions printOptions = new PrintOptions();
            printOptions.SelectionType = SelectionType.EntireFile;

            // Print Excel workbook to default printer (e.g. 'Microsoft Print to Pdf').
            string printerName = null;
            workbook.Print(printerName, printOptions);
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            print();
        }

        public void printPDFWithAcrobat()
        {
            string Filepath = @"E:\TECKSOL\Tailieu\BAOCAO.xlsx";

            using (PrintDialog Dialog = new PrintDialog())
            {
                Dialog.ShowDialog();

                ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                {
                    Verb = "print",
                    CreateNoWindow = true,
                    FileName = Filepath,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Process printProcess = new Process();
                printProcess.StartInfo = printProcessInfo;
                printProcess.Start();

                printProcess.WaitForInputIdle();

                Thread.Sleep(3000);

                if (false == printProcess.CloseMainWindow())
                {
                    printProcess.Kill();
                }
            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            printPDF();
        }
        public void printPDF()
        {
            // Absolute path to your PDF to print (with filename)
            string Filepath = @"C:\Users\DELL\Documents\Zalo Received Files\PO-210300062.pdf";
            // The name of the PDF that will be printed (just to be shown in the print queue)
            string Filename = "PO-210300062.pdf";
            // The name of the printer that you want to use
            // Note: Check step 1 from the B alternative to see how to list
            // the names of all the available printers with C#
            string PrinterName = "OneNote for Windows 10";

            // Create an instance of the Printer
            IPrinter printer = new Printer();

            // Print the file
            printer.PrintRawFile(PrinterName, Filepath, Filename);
        }
    }
}
