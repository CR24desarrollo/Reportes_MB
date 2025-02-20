using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using Reportes_MyBussines.Modelos;
using OfficeOpenXml;
using System.IO;
using Telerik.WinControls.VirtualKeyboard;
using static Microsoft.IO.RecyclableMemoryStreamManager;


namespace Reportes_MyBussines.Funciones
{
    public class Reporte3
    {
        public static event Action<int> _ProgressChanged; // Evento para reportar progreso
        public static event Action<string> _MessageUpdated; // Evento para enviar mensajes de progreso
        public static event Action _ProcessingCompleted; // Evento para indicar que se ha completado el proceso
        public static event Action<string> _Error;

        public Reporte3(Action<int> ProgressChanged, Action<string> MessageUpdated, Action ProcessingCompleted, Action<string> Error)
        {
            _ProgressChanged = ProgressChanged;
            _MessageUpdated = MessageUpdated;
            _ProcessingCompleted = ProcessingCompleted;
            _Error = Error;
        }


        public static void CrearReporte(string Fecha_Inicio, string Fecha_Final)
        {
            List<Reporte3_modelo> lsInfo = new List<Reporte3_modelo>();
            string query = $@"SELECT CONVERT(VARCHAR, F_emision, 103) AS Fecha, CONCAT(serieDocumento,NO_REFEREN) AS FACTURA, ESTADO AS ESTATUS  from ventas WHERE tipo_doc='FAC' AND F_emision >= @Inicio AND F_emision <= @Fin ";
            // Habilitar el uso de EPPlus sin licencia comercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Buscar la carpeta de descargas
            string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
            //Contadores para la barra de progreso
            int iRow = 0, rowCount = 0;

            rowCount = ObtenerTotalRegistros(Fecha_Inicio, Fecha_Final);

            try
            {
                using (SqlConnection conec = new SqlConnection(ManejoDatos.BDacceso()))
                {
                    conec.Open();

                    SqlCommand cmd = new SqlCommand(query, conec);
                    cmd.Parameters.AddWithValue("@Inicio", Fecha_Inicio);
                    cmd.Parameters.AddWithValue("@Fin", Fecha_Final);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                lsInfo.Add(new Reporte3_modelo()
                                {
                                    Fecha = reader["Fecha"].ToString(),
                                    FACTURA = reader["FACTURA"].ToString(),
                                    ESTATUS = reader["ESTATUS"].ToString()
                                });
                                iRow++;


                                _ProgressChanged?.Invoke((iRow) * 100 / rowCount);
                                _MessageUpdated?.Invoke($"Progreso: {(iRow - 1) * 100 / rowCount}%.");
                            }
                            // Exportar a Excel
                            ExportarAExcel(lsInfo, downloadsPath + $@"\\ESTATUS FACTURAS - {Fecha_Inicio} al {Fecha_Final}.xlsx");
                            _ProcessingCompleted?.Invoke();
                            MessageBox.Show("Se realizo el archivo", "[ALERT]");
                        }
                        else
                        {
                            MessageBox.Show("No se encontrarón datos en el rango de fechas.");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ManejoDatos.Log("Reporte3->CrearReporte", "[ERROR]", ex.ToString());
                MessageBox.Show("[ERROR]: " + ex);
            }
        }


        static void ExportarAExcel<T>(List<T> lista, string rutaArchivo)
        {
            using (var package = new ExcelPackage())
            {
                var hoja = package.Workbook.Worksheets.Add("Datos");

                // Convertir lista a una tabla en Excel
                hoja.Cells["A1"].LoadFromCollection(lista, true);

                // Guardar archivo
                File.WriteAllBytes(rutaArchivo, package.GetAsByteArray());
            }
        }

        private static int ObtenerTotalRegistros(string Fecha_Inicio, string Fecha_Final)
        {
            try
            {
                string query = $@"SELECT SELECT COUNT(*) FROM ventas WHERE tipo_doc = 'FAC' AND F_emision >= @Inicio AND F_emision <= @Fin";
                int total = 0;

                using (SqlConnection connection = new SqlConnection(ManejoDatos.BDacceso()))
                {
                    connection.Open();
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@Inicio", Fecha_Inicio);
                        cmd.Parameters.AddWithValue("@Fin", Fecha_Final);
                        total = (int)cmd.ExecuteScalar();
                    }
                }
                return total;
            }
            catch (Exception ex)
            {
                ManejoDatos.Log("Reporte3->ObtenerTotalRegistros", "[ERROR]", ex.ToString());
                return 9999;
            }


        }
    }
}
