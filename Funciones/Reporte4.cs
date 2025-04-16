using OfficeOpenXml;
using Reportes_MyBussines.Modelos;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reportes_MyBussines.Funciones
{
    class Reporte4
    {
        public static event Action<int> _ProgressChanged; // Evento para reportar progreso
        public static event Action<string> _MessageUpdated; // Evento para enviar mensajes de progreso
        public static event Action _ProcessingCompleted; // Evento para indicar que se ha completado el proceso
        public static event Action<string> _Error;

        public Reporte4(Action<int> ProgressChanged, Action<string> MessageUpdated, Action ProcessingCompleted, Action<string> Error)
        {
            _ProgressChanged = ProgressChanged;
            _MessageUpdated = MessageUpdated;
            _ProcessingCompleted = ProcessingCompleted;
            _Error = Error;
        }


        public static void CrearReporte(string Fecha_Inicio, string Fecha_Final)
        {
            List<Reporte4_modelo> lsInfo = new List<Reporte4_modelo>();
            string query = 
                $@"SELECT v.venta AS VtaId, CONVERT(VARCHAR, v.f_emision, 103) As Fecha, c.CP, c.Zona, IsNull(dl.LocNom,'') AS Localidad, IsNull(de.EdoNom,'') AS Estado,

                   ltrim(ltrim(c.cliente)) AS Cliente, c.nombre As Nombre, v.vend AS Vendedor, v.tipo_doc AS Tipo, v.seriedocumento as Serie, v.no_referen AS Documento,
       
                   CONCAT(REPLACE(vc.Seriedocumento, ' ', ''), REPLACE(vc.No_referen, ' ', '')) AS Cotizacion,

                   IsNull(p.linea,'') As Laboratorio, pv.articulo As Codigo, p.descrip As Articulo,

                   (pv.cantidad-pv.devconf) As Piezas, Round(pv.precio,2) As Precio,

                   Round((pv.precio*(pv.cantidad-pv.devconf)),2) As Importe,

                   Round(((pv.precio*(pv.cantidad-pv.devconf))*(pv.impuesto/100)),2) As Impuesto,

                   Round(((pv.precio*(pv.cantidad-pv.devconf))*(1+(pv.impuesto/100))),2) As Total

                   FROM partvta pv INNER JOIN prods p ON p.articulo = pv.articulo

                   INNER JOIN ventas v ON v.venta = pv.venta INNER JOIN clients c ON c.cliente = v.cliente

                   LEFT JOIN Dir_Sat_Edo de ON de.EdoCod = c.Estado

                   LEFT JOIN Dir_Sat_Loc dl ON dl.EdoCod = c.Estado AND dl.LocCod = c.Localidad
       
                   LEFT JOIN ventas vc ON vc.VentaOrigen = v.venta AND vc.tipo_doc = 'COT'

                   WHERE v.estado = 'CO' AND v.tipo_doc IN ('REM','FAC') AND IsNUll(v.cierre,0) = 0

                   AND p.articulo NOT IN('SYS','')

                   AND v.F_Emision >= @Inicio AND v.F_Emision <= @Fin 
       
                ORDER BY c.CP, c.Zona, v.Cliente, v.venta, pv.Articulo";
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
                                lsInfo.Add(new Reporte4_modelo()
                                {
                                    FECHA = reader["Fecha"].ToString(),
                                    CP = reader["CP"].ToString(),
                                    ZONA = reader["Zona"].ToString(),
                                    CIUDAD = reader["Localidad"].ToString(),
                                    ESTADO = reader["Estado"].ToString(),
                                    CLIENTE = reader["Cliente"].ToString(),
                                    NOMBRE = reader["Nombre"].ToString(),
                                    VENDEDOR = reader["Vendedor"].ToString(),
                                    TIPO = reader["Tipo"].ToString(),
                                    SERIE = reader["Serie"].ToString(),
                                    DOCUMENTO = reader["Documento"].ToString(),
                                    COTIZACION = reader["Cotizacion"].ToString(),
                                    LABORATORIO = reader["Laboratorio"].ToString(),
                                    ARTICULO = reader["Codigo"].ToString(),
                                    DESCRIPCION = reader["Articulo"].ToString(),
                                    CANTIDAD = reader["Piezas"].ToString(),
                                    PRECIO = reader["Precio"].ToString(),
                                    IMPORTE = reader["Importe"].ToString(),
                                    IMPUESTO = reader["Impuesto"].ToString(),
                                    TOTAL = reader["Total"].ToString()
                                });
                                iRow++;

                                //Mensaje y coloreado de la barra de avance
                                _ProgressChanged?.Invoke((iRow) * 100 / rowCount);
                                _MessageUpdated?.Invoke($"Progreso: {(iRow - 1) * 100 / rowCount}%.");
                            }
                            // Exportar a Excel
                            ExportarAExcel(lsInfo, downloadsPath + $@"\\VENTAS POR CP - {Fecha_Inicio} al {Fecha_Final}.xlsx");

                            _ProcessingCompleted?.Invoke();//Proceso para cambiar estatus

                            //Mesnaje de alerta
                            MessageBox.Show($@"Operación completada con éxito.{Environment.NewLine}El archivo se guardo en DESCARGAS con el Nombre: VENTAS POR CP - {Fecha_Inicio} al {Fecha_Final}.xlsx", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            //Warning
                            MessageBox.Show("No se encontrarón datos en el rango de fechas.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                _Error?.Invoke($"Error: {ex.ToString()}");
                ManejoDatos.Log("Reporte4->CrearReporte", "[ERROR]", ex.ToString());
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
                string query = 
                    $@"SELECT COUNT(v.venta) 

                               FROM partvta pv INNER JOIN prods p ON p.articulo = pv.articulo

                               INNER JOIN ventas v ON v.venta = pv.venta INNER JOIN clients c ON c.cliente = v.cliente

                               LEFT JOIN Dir_Sat_Edo de ON de.EdoCod = c.Estado

                               LEFT JOIN Dir_Sat_Loc dl ON dl.EdoCod = c.Estado AND dl.LocCod = c.Localidad
       
                               LEFT JOIN ventas vc ON vc.VentaOrigen = v.venta AND vc.tipo_doc = 'COT'

                               WHERE v.estado = 'CO' AND v.tipo_doc IN ('REM','FAC') AND IsNUll(v.cierre,0) = 0

                               AND p.articulo NOT IN('SYS','') AND v.F_emision >= @Inicio AND v.F_emision <= @Fin";
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
                ManejoDatos.Log("Reporte4->ObtenerTotalRegistros", "[ERROR]", ex.ToString());
                return 9999;
            }


        }
    }
}
