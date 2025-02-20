using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reportes_MyBussines.Funciones
{
    class ManejoDatos
    {
        public static string BDacceso()
        {
            return "Data Source=" + Properties.Settings.Default.server_name +
                                 ";Initial Catalog=" + Properties.Settings.Default.BD +
                                 ";User ID=" + Properties.Settings.Default.user_BD +
                                 ";Password=" + Properties.Settings.Default.password_BD +
                                 ";Encrypt=True; persist security info=True; trustservercertificate=True;";
        }
        public static string Log(string Funcion, string TipDoc, string exception)
        {
            string resultado = "";
            string carpeta = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            carpeta = carpeta.Replace("file:\\", "");
            carpeta = carpeta.Replace("\\bin", "\\Logs");
            try
            {
                //Si no existe la carpeta se crea
                if (!Directory.Exists(carpeta))
                    Directory.CreateDirectory(carpeta);

                //Escribir el archivo
                string rutaFinal = carpeta + "\\" + "Log-" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                File.AppendAllText(rutaFinal, Environment.NewLine + TipDoc + " - " + Funcion + ": " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " | " + exception + Environment.NewLine);

                if (!File.Exists(rutaFinal))
                    resultado = "Error al crear el archivo";
            }
            catch (Exception ex)
            {
                resultado = "Excepcion al escribir el log: " + ex.ToString();
            }
            return resultado;

        }
    }
}
