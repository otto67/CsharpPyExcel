using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace ExcelInterface
{

    public partial class Form1 : Form
    {
        private readonly string pythonFileToRun = System.IO.Directory.GetParent(System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)).Parent.FullName + 
            System.IO.Path.DirectorySeparatorChar + "RunSim.py";

        private readonly string inputFile = System.IO.Directory.GetParent(System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)).Parent.FullName +
            System.IO.Path.DirectorySeparatorChar + "input.txt";
        private readonly string resultsFile = System.IO.Directory.GetParent(System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)).Parent.FullName +
            System.IO.Path.DirectorySeparatorChar + "nodevals.txt";

        public Form1()
        {
            InitializeComponent();
        }

        private void CreateButton_Click(object sender, EventArgs e){ RunSim(); }      
        
        private void pictureBox1_Click(object sender, EventArgs e){}

        private void RunSim(){ RunSimulator(); PlotResults(); }

        private void RunSimulator()
        {
            this.outputBox.Text = "";

            string progName = @"C:\Users\otto_\winpython\python.exe";
            try
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();

                p.StartInfo = new System.Diagnostics.ProcessStartInfo(progName, pythonFileToRun)
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                p.Start();
                
                string output = p.StandardOutput.ReadToEnd();

                p.WaitForExit();

                // Displays all output from Simulator in output box
                this.outputBox.Text = output;
                
            }
            catch (Exception exc)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, exc.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, exc.Source);
            }
        }
        private void PlotResults()
        {
            Excel.Application myXL = null;
            Excel.Workbook myWB = null;
            Excel.Worksheet resultSheet = null, plotSheet = null, inputSheet = null;
            
            try
            {
                // Can be used for temporarily changing globalization.
                // System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US", true);
                
                myXL = new Excel.Application();
                myWB = (Excel.Workbook)(myXL.Workbooks.Add(Missing.Value));
                resultSheet = (Excel.Worksheet)myXL.Worksheets.Add();
                resultSheet.Name = "Result data";
                inputSheet = (Excel.Worksheet)myXL.Worksheets.Add();
                inputSheet.Name = "Input data";


                // Read results from file written by Python solver                                
                ReadResults(resultSheet);
                
                // Read input from input file
                ReadInput(inputSheet);

                CreatePlot(resultSheet);

                for (int i = myXL.ActiveWorkbook.Worksheets.Count; i > 0; i--){
                    
                    Excel.Worksheet wkSheet = (Excel.Worksheet)myXL.ActiveWorkbook.Worksheets[i];
                    if (wkSheet.Name == "Ark1" || wkSheet.Name == "Sheet1")
                        wkSheet.Delete();
                }

                myXL.Visible = true;
                myXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                if (plotSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(plotSheet);
                if (resultSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(resultSheet);
                if (inputSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(inputSheet);
                if (myWB != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(myWB);
                if (myXL != null)
                {
                    myXL.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(myXL);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void CreatePlot(Excel.Worksheet dataSheet)
        {
            Excel.ChartObjects chartObjs = null;
            Excel.ChartObject chartObj = null;
            Excel.ChartObject chartObjTV = null;
            Excel.ChartObject chartObjWF = null;
            Excel.Chart xlChart = null;
            Excel.Chart xlChartTV = null;
            Excel.Chart xlChartWF = null;
            Excel.Range chartRng = null;

            try
            {

                chartObjs = (Excel.ChartObjects)dataSheet.ChartObjects(Type.Missing);
                
                chartObj = chartObjs.Add(100, 20, 300, 300);
                chartObjWF = chartObjs.Add(100, 20, 300, 300);
                chartObjTV = chartObjs.Add(100, 20, 300, 300); // Top view
                
                xlChart = chartObj.Chart;
                xlChartTV = chartObjTV.Chart;
                xlChartWF = chartObjWF.Chart;

                // Number of data rows
                var nRows = dataSheet.Cells[1, 1].Value;

                dataSheet.Cells[1, 1] = ""; // Avoid artifacts in plot

                string upperLeftCell = "A1";
                var endRowNumber = nRows + 1;
                char endColumnLetter = (char)('A' + endRowNumber - 1);
                string lowerRightCell = $"{endColumnLetter}{endRowNumber}";


                chartRng = dataSheet.get_Range(upperLeftCell, lowerRightCell);

                xlChart.SetSourceData(chartRng, Type.Missing);
                xlChart.ChartType = Excel.XlChartType.xlSurface;
                         
                xlChartTV.SetSourceData(chartRng, Type.Missing);
                xlChartTV.ChartType = Excel.XlChartType.xlSurfaceTopView;

                xlChartWF.SetSourceData(chartRng, Type.Missing);
                xlChartWF.ChartType = Excel.XlChartType.xlSurfaceWireframe;


                // Remove axes to make the plot cleaner
                xlChart.HasAxis[Excel.XlAxisType.xlCategory,Excel.XlAxisGroup.xlPrimary] = false;
                xlChart.HasAxis[Excel.XlAxisType.xlSeriesAxis, Excel.XlAxisGroup.xlPrimary] = false;

                xlChartTV.HasAxis[Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary] = false;
                xlChartTV.HasAxis[Excel.XlAxisType.xlSeriesAxis, Excel.XlAxisGroup.xlPrimary] = false;

                xlChartWF.HasAxis[Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary] = false;
                xlChartWF.HasAxis[Excel.XlAxisType.xlSeriesAxis, Excel.XlAxisGroup.xlPrimary] = false;

             
                Excel.Axis zAxis = (Excel.Axis)xlChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                zAxis.HasTitle = true;
                zAxis.AxisTitle.Text = "f(x,y)"; 

                xlChart.HasTitle = true;
                xlChart.ChartTitle.Text = "Perspective view of solution";

                xlChart.HasLegend = false;

                // Customize axes:
                Excel.Axis xAxisTV = (Excel.Axis)xlChartTV.Axes(Excel.XlAxisType.xlCategory,
                Excel.XlAxisGroup.xlPrimary);
                xAxisTV.HasTitle = true;
                xAxisTV.AxisTitle.Text = "X Axis";
                xAxisTV.MinorTickMark = Excel.XlTickMark.xlTickMarkNone;
                                           
                
                Excel.Axis yAxisTV = (Excel.Axis)xlChartTV.Axes(Excel.XlAxisType.xlSeriesAxis, Excel.XlAxisGroup.xlPrimary);
                yAxisTV.HasTitle = true;
                yAxisTV.AxisTitle.Text = "Y Axis";
                yAxisTV.MinorTickMark = Excel.XlTickMark.xlTickMarkNone;

                xlChartTV.HasTitle = true;
                xlChartTV.ChartTitle.Text = "Top view of solution";                                
                xlChartTV.HasLegend = false;                

                xlChartTV.Location(Excel.XlChartLocation.xlLocationAsNewSheet, "Top view");
                xlChartWF.Location(Excel.XlChartLocation.xlLocationAsNewSheet, "Wire frame view");
                xlChart.Location(Excel.XlChartLocation.xlLocationAsNewSheet, "Side view");
            }
            finally
            {
                if (chartRng != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(chartRng);
                if (xlChart != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlChart);
                if (xlChartTV != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlChartTV);
                if (xlChartWF != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlChartWF);
                if (chartObj != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(chartObj);
                if (chartObjTV != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(chartObjTV);
                if (chartObjWF != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(chartObjWF);
                if (chartObjs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(chartObjs);
                              
                GC.Collect();
                GC.WaitForPendingFinalizers();                
            }
        }

        private void ReadResults(Excel.Worksheet resultSheet)
        {
            
            // Avoid trouble with wrong format floating point numbers
            System.Globalization.CultureInfo ci = (System.Globalization.CultureInfo)System.Globalization.CultureInfo.CurrentCulture.Clone();
            ci.NumberFormat.NumberDecimalSeparator = ".";
            ci.NumberFormat.CurrencyDecimalSeparator = ".";

            string[] lines = System.IO.File.ReadAllLines(resultsFile);

            for (int i = 0; i < lines.Length; i++){
                
                string[] nums = lines[i].Split(',');
                
                for (int j = 0; j < nums.Length; j++) 
                    resultSheet.Cells[i + 1, j + 1] = float.Parse(nums[j], ci);              
            }
        }

        private void ReadInput(Excel.Worksheet inputSheet)
        {
            Excel.Range myRange = null;
            
            string[] lines = System.IO.File.ReadAllLines(inputFile);

            inputSheet.Cells[1, 1]= "Input data for Poisson solver";

            myRange = inputSheet.get_Range("A1", "F14");
            myRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic);
            myRange.Interior.ColorIndex = 3;


            myRange = inputSheet.get_Range("B3", "E13");
            myRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic);
            myRange.Interior.ColorIndex = 6;

            for (int i = 0; i < lines.Length; i++){
                string[] nums = lines[i].Split(':');
                inputSheet.Cells[i + 3, 2] = nums[0];
                inputSheet.Cells[i + 3, 4] = nums[1];
            }
        }


        private void DisplayInput()
        {
            string[] lines = System.IO.File.ReadAllLines(inputFile);

            string output = "Input data for Poisson solver \r\n";
            
            foreach (var line in lines)
                output += "\r\n" + line;

            this.outputBox.Text = output;            
        }

        private void SaveInput()
        {
            System.IO.TextReader read = new System.IO.StringReader(outputBox.Text);

            string input;
            var output = new List<string>();
            
            for (int i = 0; i < outputBox.Lines.Length; i++)
            {
                input = read.ReadLine();
                if ((input != null) && (input.Split(':').Length == 2))
                    output.Add(input);
            }

            
            System.IO.File.WriteAllLines(inputFile, (string []) output.ToArray());
        }

        private void OperCombo_Click(object sender, EventArgs e){}

        private void OperCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (OperCombo.SelectedIndex < 0)
                return;

            String selected = OperCombo.SelectedItem.ToString();
            if (selected == null)
                return;

            if (selected == "Run")
                RunSim();
            else if (selected == "List input")
                DisplayInput();
            else if (selected == "Save input")
                SaveInput();
        }

        private void label1_Click(object sender, EventArgs e){}
    }
 }

