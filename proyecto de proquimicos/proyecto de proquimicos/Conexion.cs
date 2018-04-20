using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Windows.Forms;
using proyecto_de_proquimicos.Properties;
using System.Configuration;

namespace proyecto_de_proquimicos
{
    class Conexion
    {
        SqlConnection objConexion;
        SqlConnection objConexion2;
        SqlCommand command;
        SqlDataAdapter da;
        DataTable dt;

        //metodo que hace referencia con el archivo App.config
        public static string ObtenerString()
        {
            return Settings.Default.Ventas1ConnectionString;
        }

        public static string ObtenerStringProcedimiento()
        {
            return Settings.Default.ProquimicosConnectionString;
        }

        public Conexion()
        {
            objConexion = new SqlConnection(ObtenerString());
        }
        

        //hacer la conexion con la base de datos
        //public Conexion(string fe)
        //{
        //    try
        //    {
        //        //realizar la conexion
        //        objConexion2 = new SqlConnection(ObtenerStringProcedimiento());
        //        objConexion2.Open();

        //        //ejecutar el procedimiento de almacenado
        //        command = new SqlCommand("PROQUIMICOS.dbo.ESP_INFORMES_NUBE", objConexion2);
        //        command.CommandType = CommandType.StoredProcedure;

        //        SqlParameter paramCodRetorno = new SqlParameter(fe, 1);
        //        paramCodRetorno.Direction = ParameterDirection.Output;
        //        command.Parameters.Add(paramCodRetorno);

        //        command.ExecuteNonQuery();

        //        //MessageBox.Show("se conecto a la base de datos ");
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show("no se conecto con la bases de datos" + e.ToString());
        //    }
        //}

        public void CargarDatagridview(DataGridView dgv)
        {
            try
            {
                //INFORMES.DBO.VENTAS_Y_COSTOS
                da = new SqlDataAdapter("Select * from Producto", objConexion);
                dt = new DataTable();
                da.Fill(dt);
                dgv.DataSource = dt;



            }
            catch (Exception ex)
            {
                //MessageBox.Show("No se pudo  llenar  el DataGridview: " + ex.ToString());
            }

        }

        public void CargarDatagridview2(DataGridView dgv)
        {
            try
            {
                da = new SqlDataAdapter("Select * from TIEMPO_PAGO_PROVEE ", objConexion);
                dt = new DataTable();
                da.Fill(dt);
                dgv.DataSource = dt;



            }
            catch (Exception ex)
            {
               // MessageBox.Show("No se pudo  llenar  el DataGridview: " + ex.ToString());
            }

        }

        public void CargarDatagridview3(DataGridView dgv)
        {
            try
            {
                da = new SqlDataAdapter("Select * from INFORMES.DBO.TABLA_COBROS ", objConexion);
                dt = new DataTable();
                da.Fill(dt);
                dgv.DataSource = dt;



            }
            catch (Exception ex)
            {
               // MessageBox.Show("No se pudo  llenar  el DataGridview: " + ex.ToString());
            }

        }

        public void CargarDatagridview4(DataGridView dgv)
        {
            try
            {
                da = new SqlDataAdapter("Select * from INFORMES.DBO.PROD_TERMINADO ", objConexion);
                dt = new DataTable();
                da.Fill(dt);
                dgv.DataSource = dt;
            }
            catch (Exception ex)
            {
             // MessageBox.Show("No se pudo  llenar  el DataGridview: " + ex.ToString());
            }

        }

        public void CargarDatagridview5(DataGridView dgv)
        {
            try
            {
                da = new SqlDataAdapter("Select * from INFORMES.DBO.MAT_PRIMA ", objConexion);
                dt = new DataTable();
                da.Fill(dt);
                dgv.DataSource = dt;

            }
            catch (Exception ex)
            {
              //  MessageBox.Show("No se pudo  llenar  el DataGridview: " + ex.ToString());
            }

        }

        public void CargarDatagridview6(DataGridView dgv)
        {
            try
            {
                da = new SqlDataAdapter("Select * from INFORMES.DBO.CXC ", objConexion);
                dt = new DataTable();
                da.Fill(dt);
                dgv.DataSource = dt;

            }
            catch (Exception ex)
            {
             //   MessageBox.Show("No se pudo  llenar  el DataGridview: " + ex.ToString());
            }

        }


        //public static void cerrarConexion()
        //{

        //    if (objConexion != null)
        //    {
        //        objConexion.Close();
        //    }
        //}
    }
}

