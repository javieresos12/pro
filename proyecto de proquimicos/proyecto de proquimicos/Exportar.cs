using System;
using System.Windows.Forms;

namespace proyecto_de_proquimicos
{
    class Exportar
    {
        public object ActiveWorkbook { get; private set; }



        //Exportar Datagridviwe a Archivo de excel 
        public void ExportarDatosgridViewExcel( DataGridView grd)
         {

             try
             {
                    /*SaveFileDialog fichero = new SaveFileDialog();
                    fichero.Filter = "Excel ('.xls)|'.xls ";
                    fichero.FileName = "VENTAS_Y_COSTOS";
                if (fichero.ShowDialog() == DialogResult.OK)
                {*/
                     Microsoft.Office.Interop.Excel.Application aplicacion;
                     Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                     Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                     aplicacion = new Microsoft.Office.Interop.Excel.Application();
                     libros_trabajo = aplicacion.Workbooks.Add();
                     hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                    
                    for (int i = 0; i < grd.ColumnCount; i++)
                    {
                        hoja_trabajo.Cells[1,i+1] = grd.Columns[i].Name;
                    }


                   //Recorremos el DataGridview rellenando la hoja de trabajo 
                    for (int i = 0; i <= grd.Rows.Count-1; i++)
                     {
                         for (int j = 0; j < grd.Columns.Count; j++)
                         {
                             if ((grd.Rows[i].Cells[j].Value == null) == false)
                             {
                                 hoja_trabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                             }
                         }
                     }
                
               libros_trabajo.SaveAs(@"C:\Users\Javier Escobar\Downloads\VENTAS_Y_COSTOS.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
               libros_trabajo.Close(true);
               aplicacion.Quit();
                }

            //}
             catch(Exception ex)
             {
                // MessageBox.Show("Error al exportar la indormacion debido a: " + ex.ToString());
             }

         }


        public void ExportarDatosgridViewExcel2(DataGridView grd)
        {

            try
            {
                //SaveFileDialog fichero = new SaveFileDialog();
                //fichero.Filter = "Excel ('.xls)|'.xls ";
                //fichero.FileName = "TIEMPO_PAGO_PROVEE";
                //if (fichero.ShowDialog() == DialogResult.OK)
                //{
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajo = aplicacion.Workbooks.Add();
                    hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                    //Recorremos el DataGridview rellenando la hoja de trabajo 

                    for (int i = 0; i < grd.ColumnCount; i++)
                    {
                        hoja_trabajo.Cells[1, i + 1] = grd.Columns[i].Name;
                    }

                    for (int i = 0; i <= grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if ((grd.Rows[i].Cells[j].Value == null) == false)
                            {
                                hoja_trabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }


                    libros_trabajo.SaveAs(@"C:\Users\Javier Escobar\Downloads\TIEMPO_PAGO_PROVEE.xls",
                         Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajo.Close(true);
                    aplicacion.Quit();
                //}

            }
            catch (Exception ex)
            {
              //  MessageBox.Show("Error al exportar la indormacion debido a: " + ex.ToString());
            }

        }


        public void ExportarDatosgridViewExcel3(DataGridView grd)
        {

            try
            {
                //SaveFileDialog fichero = new SaveFileDialog();
                //fichero.Filter = "Excel ('.xls)|'.xls ";
                //fichero.FileName = "TABLA_COBROS";
                //if (fichero.ShowDialog() == DialogResult.OK)
                //{
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajo = aplicacion.Workbooks.Add();
                    hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                    //Recorremos el DataGridview rellenando la hoja de trabajo 

                    for (int i = 0; i < grd.ColumnCount; i++)
                    {
                        hoja_trabajo.Cells[1, i + 1] = grd.Columns[i].Name;
                    }

                    for (int i = 0; i <= grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if ((grd.Rows[i].Cells[j].Value == null) == false)
                            {
                                hoja_trabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }


                    libros_trabajo.SaveAs(@"C:\Users\Javier Escobar\Downloads\TABLA_COBROS.xls",
                         Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajo.Close(true);
                    aplicacion.Quit();
                //}

            }
            catch (Exception ex)
            {
              //  MessageBox.Show("Error al exportar la indormacion debido a: " + ex.ToString());
            }

        }


        public void ExportarDatosgridViewExcel4(DataGridView grd)
        {

            try
            {
                //SaveFileDialog fichero = new SaveFileDialog();
                //fichero.Filter = "Excel ('.xls)|'.xls ";
                //fichero.FileName = "PROD_TERMINADO";
                //if (fichero.ShowDialog() == DialogResult.OK)
                //{
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajo = aplicacion.Workbooks.Add();
                    hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                    //Recorremos el DataGridview rellenando la hoja de trabajo 

                    for (int i = 0; i < grd.ColumnCount; i++)
                    {
                        hoja_trabajo.Cells[1, i + 1] = grd.Columns[i].Name;
                    }

                    for (int i = 0; i <= grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if ((grd.Rows[i].Cells[j].Value == null) == false)
                            {
                                hoja_trabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }


                    libros_trabajo.SaveAs(@"C:\Users\Javier Escobar\Downloads\PROD_TERMINADO.xls",
                         Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajo.Close(true);
                    aplicacion.Quit();
                //}

            }
            catch (Exception ex)
            {
            //    MessageBox.Show("Error al exportar la indormacion debido a: " + ex.ToString());
            }

        }


        public void ExportarDatosgridViewExcel5(DataGridView grd)
        {

            try
            {
                //SaveFileDialog fichero = new SaveFileDialog();
                //fichero.Filter = "Excel ('.xls)|'.xls ";
                //fichero.FileName = "MAT_PRIMA";
                //if (fichero.ShowDialog() == DialogResult.OK)
                //{
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajo = aplicacion.Workbooks.Add();
                    hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                    //Recorremos el DataGridview rellenando la hoja de trabajo 

                    for (int i = 0; i < grd.ColumnCount; i++)
                    {
                        hoja_trabajo.Cells[1, i + 1] = grd.Columns[i].Name;
                    }

                    for (int i = 0; i <= grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if ((grd.Rows[i].Cells[j].Value == null) == false)
                            {
                                hoja_trabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }


                    libros_trabajo.SaveAs(@"C:\Users\Javier Escobar\Downloads\MAT_PRIMA.xls",
                         Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajo.Close(true);
                    aplicacion.Quit();
                //}

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error al exportar la indormacion debido a: " + ex.ToString());
            }

        }

        public void ExportarDatosgridViewExcel6(DataGridView grd)
        {

            try
            {
                //SaveFileDialog fichero = new SaveFileDialog();
                //fichero.Filter = "Excel ('.xls)|'.xls ";
                //fichero.FileName = "CXC";
                //if (fichero.ShowDialog() == DialogResult.OK)
                //{
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajo = aplicacion.Workbooks.Add();
                    hoja_trabajo = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);
                    //Recorremos el DataGridview rellenando la hoja de trabajo 

                    for (int i = 0; i < grd.ColumnCount; i++)
                    {
                        hoja_trabajo.Cells[1, i + 1] = grd.Columns[i].Name;
                    }

                    for (int i = 0; i <= grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if ((grd.Rows[i].Cells[j].Value == null) == false)
                            {
                                hoja_trabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }


                    libros_trabajo.SaveAs(@"C:\Users\Javier Escobar\Downloads\CXC.xls",
                         Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajo.Close(true);
                    aplicacion.Quit();
                //}

            }
            catch (Exception ex)
            {
             //   MessageBox.Show("Error al exportar la indormacion debido a: " + ex.ToString());
            }

        }
    }
}
