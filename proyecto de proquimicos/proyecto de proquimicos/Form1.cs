using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace proyecto_de_proquimicos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DateTime fe = DateTime.Now;
            string actual = fe.ToShortDateString();
            //Conexion c2 = new Conexion(actual);
            Conexion c = new Conexion();
            Exportar exp = new Exportar();
            for (int i = 1; i <= 6; i++)
            {
                if (i == 1)
                {
                 c.CargarDatagridview(Grilla);
                 exp.ExportarDatosgridViewExcel(Grilla);
                }
                else if (i == 2)
                {
                    c.CargarDatagridview2(Grilla);
                    exp.ExportarDatosgridViewExcel2(Grilla);
                }
                else if (i == 3)
                {
                    c.CargarDatagridview3(Grilla);
                    exp.ExportarDatosgridViewExcel3(Grilla);
                }
                else if (i == 4)
                {
                    c.CargarDatagridview4(Grilla);
                    exp.ExportarDatosgridViewExcel4(Grilla);
                }
                else if (i == 5)
                {
                    c.CargarDatagridview5(Grilla);
                    exp.ExportarDatosgridViewExcel5(Grilla);
                }
                else if (i == 6)
                {
                    c.CargarDatagridview6(Grilla);
                    exp.ExportarDatosgridViewExcel6(Grilla);
                }

            }

        }

        private void label_Click(object sender, EventArgs e)
        {

        }

        private void Grilla_CellContentClick(object sender, DataGridViewCellEventArgs grilla)
        {
         
        }

        private void dataGridView1_DefaultValuesNeeded(object sender,
             System.Windows.Forms.DataGridViewRowEventArgs e)
        {

        }

        private void cmbTabla_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
