﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class Form1 : Form
    {
        List<Flat> flats;
        RealEstateEntities re = new RealEstateEntities();


        void LoadData()
        {
            flats = re.Flats.ToList();
        }

        public Form1()
        {
            InitializeComponent();
            LoadData();
        }
    }
}
