using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Value_at_Risk.Entities;

namespace Value_at_Risk
{
    public partial class Form1 : Form
    {
        List<Tick> ticks;
        PortfolioEntities context = new PortfolioEntities();
        List<PortfolioItem> Portfolio = new List<PortfolioItem>();

        public Form1()
        {
            InitializeComponent();
            ticks = context.Tick.ToList();
            dataGridView1.DataSource = ticks;
            CreatePorfolio();
        }

        private void CreatePorfolio()
        {
            Portfolio.Add(new PortfolioItem() { Index = "OTP", Volume = 10 });
            Portfolio.Add(new PortfolioItem() { Index = "ZWACK", Volume = 10 });
            Portfolio.Add(new PortfolioItem() { Index = "ELMU", Volume = 10 });
            dataGridView2.DataSource = Portfolio;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
    }
}
