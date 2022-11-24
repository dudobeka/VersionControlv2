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
    public partial class Form1 : Form
    {
        List<Flat> flats;
        RealEstateEntities re = new RealEstateEntities();

        Excel1.Application xlApp; // A Microsoft Excel alkalmazás
        Excel1.Workbook xlWB; // A létrehozott munkafüzet
        Excel1.Worksheet xlSheet; // Munkalap a munkafüzeten belül


        void LoadData()
        {
            flats = re.Flats.ToList();
        }

        void CreateExcel()
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

        private void CreateTable()
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




        }


        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();
        }
    }
}
