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

namespace Reportes_MyBussines.Funciones
{
    public class Reporte1
    {
        public static event Action<int> _ProgressChanged; // Evento para reportar progreso
        public static event Action<string> _MessageUpdated; // Evento para enviar mensajes de progreso
        public static event Action _ProcessingCompleted; // Evento para indicar que se ha completado el proceso
        public static event Action<string> _Error;
        public Reporte1(Action<int> ProgressChanged, Action<string> MessageUpdated, Action ProcessingCompleted, Action<string> Error)
        {
            _ProgressChanged = ProgressChanged;
            _MessageUpdated = MessageUpdated;
            _ProcessingCompleted = ProcessingCompleted;
            _Error = Error;
        }
        public static void CrearReporte(string Fecha_Inicio, string Fecha_Final)
        {
            List<Reporte1_modelo> lsInfo = new List<Reporte1_modelo>();
            string query = $@"SELECT * FROM ( SELECT 
								CONVERT(VARCHAR, v.F_emision, 103) as Fecha,
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN c.CP ELSE cu.CodigoPostal  END ) AS CP,
								CONCAT(( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN c.Calle ELSE cu.Calle END )+' ',
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN c.numeroexterior ELSE cu.NumExterior END )+', ',
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN c.numerointerior ELSE cu.NumInterior END )+', ',
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN IsNull(dcc.ColNom,'') ELSE IsNull(dcu.ColNom,'')  END )+', ',
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN IsNull(dmc.MunNom,'') ELSE IsNull(dmu.MunNom,'')  END )+', ',
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN IsNull(det.EdoNom,'') ELSE IsNull(deu.EdoNom,'')  END )
								) As Dirección,
								( CASE WHEN IsNull(cu.IDor_de,'') = '' THEN c.Cliente ELSE cu.clienteId  END ) AS Codigo_Cliente,
								c.NOMBRE AS Razón_Social,
								v.TIPO_DOC AS Tipo, 
								CONCAT(v.seriedocumento,v.NO_REFEREN) AS Folio,
								pv.Articulo AS Articulo,
								pv.Observ AS Descripcion,
								p.LINEA AS Laboratorio,
								pv.Precio AS Precio,
								pv.Cantidad AS Cantidad, 
								(pv.PRECIO*pv.Cantidad) AS Importe,
								v.IMPUESTO AS Impuesto,
								v.IMPORTE AS Total,
								v.ESTADO AS ESTATUS
								FROM Partvta pv INNER JOIN prods p ON p.articulo = pv.articulo 
								INNER JOIN Ventas v ON v.Venta = pv.Venta INNER JOIN Clients c On c.Cliente = v.Cliente 
								LEFT JOIN Vends ve ON ve.Vend = v.Vend
								LEFT JOIN CPUbicaciones cu ON cu.Cliente = v.Cliente AND cu.IDor_de = v.DateMbId
								LEFT JOIN Dir_Sat_Pais dpc On dpc.PaisCod = c.Pais
								LEFT JOIN Dir_Sat_Pais dpu On dpu.PaisCod = cu.Pais
								LEFT JOIN Dir_Sat_Edo det ON det.EdoCod = c.Estado AND det.Pais = c.Pais
								LEFT JOIN Dir_Sat_Edo deu ON deu.EdoCod = cu.Estado AND deu.Pais = cu.Pais
								LEFT JOIN Dir_Sat_Mun dmc ON dmc.EdoCod = c.Estado AND dmc.MunCod = c.Pobla
								LEFT JOIN Dir_Sat_Mun dmu ON dmu.EdoCod = cu.Estado AND dmu.MunCod = cu.Municipio
								LEFT JOIN Dir_Sat_Loc dlc ON dlc.EdoCod = c.Estado AND dlc.LocCod = c.Localidad
								LEFT JOIN Dir_Sat_Loc dlu ON dlu.EdoCod = cu.Estado AND dlu.LocCod = cu.Localidad
								LEFT JOIN Dir_Sat_Col dcc ON dcc.CP = c.CP AND dcc.ColCod = c.Colonia
								LEFT JOIN Dir_Sat_Col dcu ON dcu.CP = cu.CodigoPostal AND dcu.ColCod= cu.Colonia
								WHERE 
								--v.estado = 'CO' 
								--AND 
								v.tipo_doc IN('REM','FAC','COT') 
								--AND IsNUll(v.cierre,0) = 0 
								AND p.articulo NOT IN('SYS','') 
								AND v.F_emision >= @Inicio AND v.F_emision <= @Fin
								AND pv.articulo NOT IN('SYS','') ) P0 ORDER BY Codigo_Cliente";
            // Habilitar el uso de EPPlus sin licencia comercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Buscar la carpeta de descargas
            string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
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
                                lsInfo.Add(new Reporte1_modelo()
                                {
                                    Fecha = reader["Fecha"].ToString(),
                                    CP = reader["CP"].ToString(),
                                    Direccion = reader["Dirección"].ToString(),
                                    Codigo_Cliente = reader["Codigo_Cliente"].ToString(),
                                    Razon_Social = reader["Razón_Social"].ToString(),
                                    Tipo = reader["Tipo"].ToString(),
                                    Folio = reader["Folio"].ToString(),
                                    Articulo = reader["Articulo"].ToString(),
                                    Descripcion = reader["Descripcion"].ToString(),
                                    Laboratorio = reader["Laboratorio"].ToString(),
                                    Precio = reader["Precio"].ToString(),
                                    Cantidad = reader["Cantidad"].ToString(),
                                    Importe = reader["Importe"].ToString(),
                                    Impuesto = reader["Impuesto"].ToString(),
                                    Total = reader["Total"].ToString(),
                                    ESTATUS = reader["ESTATUS"].ToString()
                                });
                                iRow++;


                                _ProgressChanged?.Invoke((iRow) * 100 / rowCount);
                                _MessageUpdated?.Invoke($"Progreso: {(iRow - 1) * 100 / rowCount}%.");
                            }
                            // Exportar a Excel
                            ExportarAExcel(lsInfo, downloadsPath+$@"\\GENERAL VENTAS - {Fecha_Inicio} al {Fecha_Final}.xlsx");
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
                _Error?.Invoke($"Error: {ex.ToString()}");
                ManejoDatos.Log("Reporte1->CrearReporte", "[ERROR]", ex.ToString());
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
                string query = $@"SELECT COUNT(*) AS Total_Registros
                                    FROM (
                                        SELECT 
                                            CONVERT(VARCHAR, v.F_emision, 103) as Fecha,
                                            (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN c.CP ELSE cu.CodigoPostal END) AS CP,
                                            CONCAT(
                                                (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN c.Calle ELSE cu.Calle END) + ' ',
                                                (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN c.numeroexterior ELSE cu.NumExterior END) + ', ',
                                                (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN c.numerointerior ELSE cu.NumInterior END) + ', ',
                                                (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN IsNull(dcc.ColNom, '') ELSE IsNull(dcu.ColNom, '') END) + ', ',
                                                (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN IsNull(dmc.MunNom, '') ELSE IsNull(dmu.MunNom, '') END) + ', ',
                                                (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN IsNull(det.EdoNom, '') ELSE IsNull(deu.EdoNom, '') END)
                                            ) AS Dirección,
                                            (CASE WHEN IsNull(cu.IDor_de, '') = '' THEN c.Cliente ELSE cu.clienteId END) AS Codigo_Cliente,
                                            c.NOMBRE AS Razón_Social,
                                            v.TIPO_DOC AS Tipo, 
                                            CONCAT(v.seriedocumento, v.NO_REFEREN) AS Folio,
                                            pv.Articulo AS Articulo,
                                            pv.Observ AS Descripcion,
                                            p.LINEA AS Laboratorio,
                                            pv.Precio AS Precio,
                                            pv.Cantidad AS Cantidad, 
                                            (pv.PRECIO * pv.Cantidad) AS Importe,
                                            v.IMPUESTO AS Impuesto,
                                            v.IMPORTE AS Total,
                                            v.ESTADO AS ESTATUS
                                        FROM Partvta pv 
                                        INNER JOIN prods p ON p.articulo = pv.articulo 
                                        INNER JOIN Ventas v ON v.Venta = pv.Venta 
                                        INNER JOIN Clients c ON c.Cliente = v.Cliente 
                                        LEFT JOIN Vends ve ON ve.Vend = v.Vend
                                        LEFT JOIN CPUbicaciones cu ON cu.Cliente = v.Cliente AND cu.IDor_de = v.DateMbId
                                        LEFT JOIN Dir_Sat_Pais dpc ON dpc.PaisCod = c.Pais
                                        LEFT JOIN Dir_Sat_Pais dpu ON dpu.PaisCod = cu.Pais
                                        LEFT JOIN Dir_Sat_Edo det ON det.EdoCod = c.Estado AND det.Pais = c.Pais
                                        LEFT JOIN Dir_Sat_Edo deu ON deu.EdoCod = cu.Estado AND deu.Pais = cu.Pais
                                        LEFT JOIN Dir_Sat_Mun dmc ON dmc.EdoCod = c.Estado AND dmc.MunCod = c.Pobla
                                        LEFT JOIN Dir_Sat_Mun dmu ON dmu.EdoCod = cu.Estado AND dmu.MunCod = cu.Municipio
                                        LEFT JOIN Dir_Sat_Loc dlc ON dlc.EdoCod = c.Estado AND dlc.LocCod = c.Localidad
                                        LEFT JOIN Dir_Sat_Loc dlu ON dlu.EdoCod = cu.Estado AND dlu.LocCod = cu.Localidad
                                        LEFT JOIN Dir_Sat_Col dcc ON dcc.CP = c.CP AND dcc.ColCod = c.Colonia
                                        LEFT JOIN Dir_Sat_Col dcu ON dcu.CP = cu.CodigoPostal AND dcu.ColCod = cu.Colonia
                                        WHERE 
                                            --v.estado = 'CO' 
                                            --AND 
                                            v.tipo_doc IN ('REM', 'FAC', 'COT') 
                                            --AND IsNUll(v.cierre, 0) = 0 
                                            AND p.articulo NOT IN ('SYS', '') 
                                            AND v.F_emision >= @Inicio AND v.F_emision <= @Fin
                                            AND pv.articulo NOT IN ('SYS', '')
                                    ) AS P0;
                                    ";
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
                ManejoDatos.Log("Reporte1->ObtenerTotalRegistros", "[ERROR]", ex.ToString());
                return 9999;
            }
           
        }
    }
}
