using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Reportes_MyBussines.Funciones;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Reportes_MyBussines
{
    public partial class Form2: Form
    {

        public Form2()
        {
            InitializeComponent();
            textBox1.Text = "Reporte General Ventas";

            dateTimePicker1.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd/MMM/yyyy";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "dd/MMM/yyyy";

            //Definimos los valores maximos y minimos para la barra de progreso
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = 100;

            // Configura el BackgroundWorker
            backgroundWorker.WorkerReportsProgress = true;

            backgroundWorker.DoWork += backgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;

            //Evitar la selecion al presionar el boton que genera los reportes
            this.textBox1.TabStop = false;
            this.textBox1.GotFocus += (s, e) => this.ActiveControl = null;
            this.textBox2.TabStop = false;
            this.textBox2.GotFocus += (s, e) => this.ActiveControl = null;
            this.textBox3.TabStop = false;
            this.textBox3.GotFocus += (s, e) => this.ActiveControl = null;
            this.dateTimePicker1.TabStop = false;
            this.dateTimePicker1.GotFocus += (s, e) => this.ActiveControl = null;
            this.dateTimePicker2.TabStop = false;
            this.dateTimePicker2.GotFocus += (s, e) => this.ActiveControl = null;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            // Resetea la barra de progreso
            toolStripProgressBar1.Value = 0;
            toolStripStatusLabel1.Text = "Iniciando proceso...";
            backgroundWorker.RunWorkerAsync(); // Inicia el proceso en segundo plano

            //Deshabilitamos los botones para evitar interrupciones
            button1.Enabled = false;
            button2.Enabled = false;

            //Cambiamos el cursor a WaitCursor
            this.Cursor = Cursors.WaitCursor;

        }

        public void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string message = string.Empty;
            string fechaInicial = dateTimePicker1.Value.ToString("yyyyMMdd");
            string fechaFinal = dateTimePicker2.Value.ToString("yyyyMMdd");
            try
            {
                Reporte1 primerReporte = new Reporte1(
                    progreso =>
                    {
                        // Verifica si el proceso debe ser cancelado
                        if (backgroundWorker.CancellationPending)
                        {
                            e.Cancel = true;
                            return;
                        }

                        // Reporta el progreso
                        backgroundWorker.ReportProgress(progreso);
                    },
                            mensaje =>
                            {
                                // Envía el mensaje a la interfaz principal
                                this.Invoke(new Action(() => toolStripStatusLabel1.Text = mensaje));
                            },
                            () =>
                            {
                                this.Invoke(new Action(() =>
                                {
                                    backgroundWorker.ReportProgress(100);
                                    this.Invoke(new Action(() => toolStripStatusLabel1.Text = "El proceso ha finalizado"));

                                }));
                            },
                            error =>
                            {
                                // Envía el mensaje a la interfaz principal
                                this.Invoke(new Action(() =>
                                {
                                    e.Result = error;
                                    MessageBox.Show($"Error: {error}", "Notificación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }));
                            }

                    );
                Reporte1.CrearReporte(fechaInicial, fechaFinal);
            }
            catch (Exception ex)
            {
                e.Result = ex.Message; // Pasa el mensaje de error al RunWorkerCompleted
            }

        }
        public void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage; // Actualiza la barra de progreso
            toolStripStatusLabel1.Text = $"Progreso: {e.ProgressPercentage}%"; // Actualiza el estado
        }
        public void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                toolStripStatusLabel1.Text = $"Error: {e.Error.Message}";
                toolStripProgressBar1.Value = 0; // Resetea la barra de progreso en caso de error
            }
            else if (e.Result is string errorMessage)
            {
                toolStripStatusLabel1.Text = $"Error: {errorMessage}";
                toolStripProgressBar1.Value = 0; // Resetea la barra de progreso en caso de error
            }
            else
            {
                toolStripStatusLabel1.Text = "Proceso completado.";
            }

            //Habilitamos los botones
            button1.Enabled = true;
            button2.Enabled = true;

            //Cambiamos el cursor a Default
            this.Cursor = Cursors.Default;

        }

    }
}
