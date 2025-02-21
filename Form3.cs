using Reportes_MyBussines.Funciones;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reportes_MyBussines
{
    public partial class Form3: Form2
    {
        public Form3()
        {
            textBox1.Text = "Reporte Remisión y Factura";
            InitializeComponent();
            button1.Click -= button1_Click;  // Elimina el evento del padre
            button1.Click += button1_Click_Hijo;  // Agrega un nuevo evento

            // Configura el BackgroundWorker con los nuevos eventos.
            backgroundWorker.DoWork -= backgroundWorker_DoWork;  // Quita el evento del padre.
            backgroundWorker.DoWork += backgroundWorker_DoWork_Hijo;  // Asigna el del hijo.
            backgroundWorker.ProgressChanged -= BackgroundWorker_ProgressChanged;  // Quita el evento del padre.
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged_Hijo;  // Asigna el del hijo.
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;  // Quita el evento del padre.
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted_Hijo;  // Asigna el del hijo.


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
        private void button1_Click_Hijo(object sender, EventArgs e)
        {
            // Resetea la barra de progreso
            toolStripProgressBar1.Value = 0;
            toolStripStatusLabel1.Text = "Iniciando proceso...";
            backgroundWorker.RunWorkerAsync(); // Inicia el proceso en segundo plano 
            button1.Enabled = false;
            button2.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

        }

        private void backgroundWorker_DoWork_Hijo(object sender, DoWorkEventArgs e)
        {
            string message = string.Empty;
            string fechaInicial = dateTimePicker1.Value.ToString("yyyyMMdd");
            string fechaFinal = dateTimePicker2.Value.ToString("yyyyMMdd");
            try
            {
                Reporte2 primerReporte = new Reporte2(
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
                Reporte2.CrearReporte(fechaInicial, fechaFinal);
            }
            catch (Exception ex)
            {
                e.Result = ex.Message; // Pasa el mensaje de error al RunWorkerCompleted
            }

        }

        private void BackgroundWorker_ProgressChanged_Hijo(object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage; // Actualiza la barra de progreso
            toolStripStatusLabel1.Text = $"Progreso: {e.ProgressPercentage}%"; // Actualiza el estado
        }
        private void BackgroundWorker_RunWorkerCompleted_Hijo(object sender, RunWorkerCompletedEventArgs e)
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
            button1.Enabled = true;
            button2.Enabled = true;
            this.Cursor = Cursors.Default;
        }
    }
}
