using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel1 = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Excel
{
    public partial class Form : System.Windows.Forms.Form
    {
        List<Flat> flats;
        RealEstateEntities re = new RealEstateEntities();

        Excel1.Application xlApp;
        Excel1.Workbook xlWB;
        Excel1.Worksheet xlSheet;

        public void LoadData()
        {
            flats = re.Flats.ToList();
        }

        public void CreateExcel()
        {
            try
            {
                xlApp = new Excel1.Application();
                xlWB = xlApp.Workbooks.Add();
                xlSheet = xlWB.ActiveSheet;

                CreateTable();

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Source + '\n' + ex.Message);
                xlWB.Close(false);
                xlApp.Quit();
                xlApp = null;
                xlWB = null;
            }

        }

        public void CreateTable()
        {

            string[] headers = new string[] {
             "Kód",
             "Eladó",
             "Oldal",
             "Kerület",
             "Lift",
             "Szobák száma",
             "Alapterület (m2)",
             "Ár (mFt)",
             "Négyzetméter ár (Ft/m2)"};

            object[,] values = new object[flats.Count, headers.Length];
            int counter = 0;
            Excel1.Range r;

            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, i + 1] = headers[i];
            }


            foreach (var f in flats)
            {
                values[counter, 0] = f.Code;
                values[counter, 1] = f.Vendor;
                values[counter, 2] = f.Side;
                values[counter, 3] = f.District;
                values[counter, 4] = f.Elevator;
                values[counter, 5] = f.NumberOfRooms;
                values[counter, 6] = f.FloorArea;
                values[counter, 7] = f.Price;
                values[counter, 8] = "";
                counter++;

            }

            r = xlSheet.get_Range(GetCell(2, 1),
                       GetCell(flats.Count + 1, headers.Length));
            r.Value = values;
            r = xlSheet.get_Range(GetCell(2, 9),
                        GetCell(flats.Count + 1, 9));
            r.Value = "=1000000*" + GetCell(2, 8) + "/" + GetCell(2, 7);

            Excel1.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel1.XlLineStyle.xlContinuous, Excel1.XlBorderWeight.xlThick);

            r = xlSheet.UsedRange;
            r.BorderAround2(Excel1.XlLineStyle.xlContinuous, Excel1.XlBorderWeight.xlThick);
            r = xlSheet.get_Range(GetCell(2, 1),
                        GetCell(flats.Count + 1, 1));
            r.Font.Bold = true;
            r.Interior.Color = Color.LightYellow;
            //utolsó oszlop halványzöld háttér
            r = xlSheet.get_Range(GetCell(2, 9),
                        GetCell(flats.Count + 1, 9));
            r.Interior.Color = Color.LightGreen;
            r.NumberFormat = "0.00";
        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }
        public Form()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();
        }
    }
 }
