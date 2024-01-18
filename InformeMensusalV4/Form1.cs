using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using DataTable = System.Data.DataTable;
//using System.Data.DataTable = System.Data.System.Data.DataTable;

namespace InformeMensusalV4
{
    public partial class Form1 : Form
    {
        //System.Data.DataTable data_Table; // Especifica el espacio de nombres para evitar la ambigüedad
        DataTable data_Table;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnCargarExcel_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Archivos Excel|*.xlsx";
            openFileDialog1.Title = "Seleccionar archivo Excel";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                CargarDatosDesdeExcel(openFileDialog1.FileName);
            }
        }

        private void CargarDatosDesdeExcel(string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var ws = workbook.Worksheets.First();
                    //data_Table = ws.RangeUsed().AsTable().AsNativeSystem.Data.DataTable();
                    data_Table = ws.RangeUsed().AsTable().AsNativeDataTable();


                    dataGridViewTickets.DataSource = data_Table;

                    // Aplicar el filtro por período de tiempo
                    AplicarFiltroPorFecha();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al cargar el archivo Excel: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AplicarFiltroPorFecha()
        {
            DateTime fechaDesde = dateTimePickerDesde.Value;
            DateTime fechaHasta = dateTimePickerHasta.Value;

            //data_Table.DefaultView.RowFilter = $"[Fecha-inicio] >= #{fechaDesde:MM/dd/yyyy}# AND [Fecha-cierre] <= #{fechaHasta:MM/dd/yyyy}#";
            if (data_Table != null) // Verifica que data_Table no sea null
            {
                data_Table.DefaultView.RowFilter = $"[Fecha-inicio] >= #{fechaDesde:MM/dd/yyyy}# AND [Fecha-cierre] <= #{fechaHasta:MM/dd/yyyy}#";
            }
            else
            {
                MessageBox.Show("Error: El DataTable es nulo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenerarPDF_Click(object sender, EventArgs e)
        {
            if (data_Table == null || data_Table.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para generar el PDF.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Lógica para generar el PDF
            GenerarPDF();
        }

        private void GenerarPDF()
        {
            using (var document = new PdfDocument())
            {
                // Añadir página al PDF
                PdfPage page = document.AddPage();

                // Obtener el objeto XGraphics para dibujar en la página
                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    // Aquí puedes añadir el contenido al PDF, por ejemplo:
                    XFont font = new XFont("Arial", 12);

                    gfx.DrawString("INFOME MENSUAL", font, XBrushes.Black, new XPoint(30, 10));

                    gfx.DrawString("------------------------------------", font, XBrushes.Black, new XPoint(30, 20));


                    gfx.DrawString($"Cantidad de tickets totales: {data_Table.Rows.Count}", font, XBrushes.Black, new XPoint(30, 30));

                    gfx.DrawString($"Cantidad de tickets en trámite: {ContarTicketsEnTramite()}", font, XBrushes.Black, new XPoint(30, 60));

                    // Información de tickets por Web Service
                    var ticketsPorWebService = ObtenerTicketsPorWebService();
                    int posY = 90;

                    foreach (var ws in ticketsPorWebService)
                    {
                        gfx.DrawString($"WebService: {ws.Key}", font, XBrushes.Black, new XPoint(30, posY));
                        posY += 20;

                        foreach (var tipo in ws.Value)
                        {
                            gfx.DrawString($"   Tipo: {tipo.Key}, Cantidad: {tipo.Value}", font, XBrushes.Black, new XPoint(30, posY));
                            posY += 20;
                        }
                    }

                    // Guardar el archivo PDF
                    string pdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "InformeTickets.pdf");
                    document.Save(pdfPath);

                    // Mostrar el PDF en el WebBrowser
                    webBrowserPDF.Url = new Uri(pdfPath);
                }
            }
        }

        private int ContarTicketsEnTramite()
        {
            if (data_Table == null || data_Table.Rows.Count == 0)
                return 0;

                int count = data_Table.AsEnumerable()
                .Count(row => row.Field<string>("Estado") == "en proceso" || row.Field<string>("Estado") == "En Revisión" || row.Field<string>("Estado") == "abierto" || row.Field<string>("Estado") == "En curso" || row.Field<string>("Estado") == "En fila");

            return count;
        }


        private Dictionary<string, Dictionary<string, int>> ObtenerTicketsPorWebService()
        {
            if (data_Table == null)
            {
                return new Dictionary<string, Dictionary<string, int>>(); // retorna un diccionario vacío si data_Table es null
            }

            // agrupo los tickets por Web Service y por tipo
            var ticketsPorWebService = data_Table.AsEnumerable()
                .Where(row => !row.IsNull("Name WS") && !row.IsNull("Tipo-TKT")) 
                .GroupBy(row => row.Field<string>("Name WS"))
                .ToDictionary(
                    wsGroup => wsGroup.Key,  // nombre del Ws
                    wsGroup => wsGroup
                        .GroupBy(tipoGroup => tipoGroup.Field<string>("Tipo-TKT"))
                        .ToDictionary(
                            tipoGroup => tipoGroup.Key,  // tipo del Ticket
                            tipoGroup => tipoGroup.Count()  // cntidad de Tickets por tipo
                        )
                );

            return ticketsPorWebService;
        }


    }
}