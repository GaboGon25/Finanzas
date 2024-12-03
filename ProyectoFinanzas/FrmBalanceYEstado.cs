using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ProyectoFinanzas
{
    public partial class FrmBalanceYEstado : Form
    {
        public FrmBalanceYEstado()
        {
            InitializeComponent();
        }

        private void FrmBalanceYEstado_Load(object sender, EventArgs e)
        {
            dgvBalanceGeneral.ColumnCount = 2;
            dgvBalanceGeneral.Columns[0].Name = "Descripción";
            dgvBalanceGeneral.Columns[1].Name = "Monto (USD)";
            dgvBalanceGeneral.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvBalanceGeneral.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);

            //dgvEstadoDeResultado.ColumnCount = 2;
            //dgvEstadoDeResultado.Columns[0].Name = "Descripción";
            //dgvEstadoDeResultado.Columns[1].Name = "Monto (USD)";
            //dgvEstadoDeResultado.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            //dgvEstadoDeResultado.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            dgvEstadoDeResultado.ColumnCount = 2;
            dgvEstadoDeResultado.Columns[0].Name = "Descripción";
            dgvEstadoDeResultado.Columns[1].Name = "Monto (USD)";

            dgvEstadoDeResultado.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvEstadoDeResultado.Columns[0].FillWeight = 70;
            dgvEstadoDeResultado.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            dgvEstadoDeResultado.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Bold);
            dgvEstadoDeResultado.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;



            if (DatosFinancieros.BalanceGeneral.Count == 0)
                DatosFinancieros.BalanceGeneral = BalanceGeneral();

            if (DatosFinancieros.EstadoResultados.Count == 0)
                DatosFinancieros.EstadoResultados = estadoResultados();

            ActualizarGrids();
            CalcularRazonesFinancieras();
            CalcularRazonesActividad();

        }


        private void ActualizarGrids()
        {
            dgvBalanceGeneral.Rows.Clear();
            foreach (var item in DatosFinancieros.BalanceGeneral)
            {
                dgvBalanceGeneral.Rows.Add(item.Descripcion, item.Monto);
            }

            dgvEstadoDeResultado.Rows.Clear();
            foreach (var item in DatosFinancieros.EstadoResultados)
            {
                dgvEstadoDeResultado.Rows.Add(item.Descripcion, item.Monto);
            }
        }


        private List<DatosBalance> BalanceGeneral()
        {
            return new List<DatosBalance>
            {
                new DatosBalance { Descripcion = "ACTIVOS CORRIENTES", Monto = "" },
                new DatosBalance { Descripcion = "Efectivo y equivalentes de efectivo", Monto = "1,086" },
                new DatosBalance { Descripcion = "Cuentas por cobrar, neto", Monto = "821" },
                new DatosBalance { Descripcion = "Inventarios", Monto = "559" },
                new DatosBalance { Descripcion = "Impuestos por recuperar", Monto = "401" },
                new DatosBalance { Descripcion = "Otros activos financieros corrientes", Monto = "57" },
                new DatosBalance { Descripcion = "Otros activos corrientes", Monto = "87" },
                new DatosBalance { Descripcion = "Total activos corrientes", Monto = "3,011" },
                new DatosBalance { Descripcion = "", Monto = "" },
                new DatosBalance { Descripcion = "Activos no corrientes", Monto = "" },
                new DatosBalance { Descripcion = "Inversiones en otras entidades", Monto = "517" },
                new DatosBalance { Descripcion = "Activos por derecho de uso", Monto = "73" },
                new DatosBalance { Descripcion = "Propiedad, planta y equipo, neto", Monto = "3,244" },
                new DatosBalance { Descripcion = "Activos intangibles, neto", Monto = "13,591" },
                new DatosBalance { Descripcion = "Activos por impuestos diferidos", Monto = "553" },
                new DatosBalance { Descripcion = "Otros activos financieros no corrientes", Monto = "11" },
                new DatosBalance { Descripcion = "Otros activos no corrientes", Monto = "321" },
                new DatosBalance { Descripcion = "Total activos no corrientes", Monto = "10,660" },
                new DatosBalance { Descripcion = "TOTAL ACTIVOS", Monto = "13,671" },
                new DatosBalance { Descripcion = "", Monto = "" },
                new DatosBalance { Descripcion = "PASIVOS Y CAPITAL", Monto = "" },
                new DatosBalance { Descripcion = "Pasivos corrientes", Monto = "" },
                new DatosBalance { Descripcion = "Préstamos bancarios y documentos por pagar", Monto = "47" },
                new DatosBalance { Descripcion = "Vencimiento a corto plazo del pasivo no corriente", Monto = "562" },
                new DatosBalance { Descripcion = "Vencimiento a corto plazo del pasivo por arrendamiento", Monto = "26" },
                new DatosBalance { Descripcion = "Intereses por pagar", Monto = "1,052" },
                new DatosBalance { Descripcion = "Proveedores", Monto = "548" },
                new DatosBalance { Descripcion = "Impuestos por pagar", Monto = "379" },
                new DatosBalance { Descripcion = "Otros pasivos financieros corrientes", Monto = "68" },
                new DatosBalance { Descripcion = "Total pasivos corrientes", Monto = "2,705" },
                new DatosBalance { Descripcion = "", Monto = "" },
                new DatosBalance { Descripcion = "Pasivos no corrientes", Monto = "" },
                new DatosBalance { Descripcion = "Préstamos bancarios y documentos por pagar", Monto = "3,101" },
                new DatosBalance { Descripcion = "Beneficios posteriores al retiro y otros beneficios a empleados", Monto = "1,135" },
                new DatosBalance { Descripcion = "Porción no corriente de pasivo por arrendamiento", Monto = "11" },
                new DatosBalance { Descripcion = "Otros pasivos no corrientes", Monto = "26" },
                new DatosBalance { Descripcion = "Provisiones y otros pasivos no corrientes", Monto = "1,602" },
                new DatosBalance { Descripcion = "Total pasivos no corrientes", Monto = "4,090" },
                new DatosBalance { Descripcion = "Total pasivos", Monto = "6,795" },
                new DatosBalance { Descripcion = "", Monto = "" },
                new DatosBalance { Descripcion = "CAPITAL", Monto = "" },
                new DatosBalance { Descripcion = "Capital Social", Monto = "109" },
                new DatosBalance { Descripcion = "Prima en suscripción de acciones", Monto = "2,416" },
                new DatosBalance { Descripcion = "Utilidades retenidas", Monto = "4,092" },
                new DatosBalance { Descripcion = "Otros ingresos de capital", Monto = "65" },
                new DatosBalance { Descripcion = "Capital atribuible a los propietarios de la controladora", Monto = "6,158" },
                new DatosBalance { Descripcion = "Participación no controladora en subsidiarias consolidadas", Monto = "718" },
                new DatosBalance { Descripcion = "Total capital", Monto = "6,876" },
                new DatosBalance { Descripcion = "TOTAL PASIVOS Y CAPITAL", Monto = "13,671" }
            };
        }

        private void CalcularRazonesFinancieras()
        {
            // Variables para almacenar valores importantes
            decimal activoCirculante = 0;
            decimal pasivoCirculante = 0;
            decimal inventarios = 0;

            // Buscar los valores en la lista BalanceGeneral
            foreach (var item in DatosFinancieros.BalanceGeneral)
            {
                switch (item.Descripcion)
                {
                    case "Total activos corrientes":
                        activoCirculante = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Total pasivos corrientes":
                        pasivoCirculante = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Inventarios":
                        inventarios = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                }
            }

            // Cálculos
            decimal capitalDeTrabajo = activoCirculante - pasivoCirculante;
            decimal razonCirculante = pasivoCirculante != 0 ? activoCirculante / pasivoCirculante : 0;
            decimal razonRapida = pasivoCirculante != 0 ? (activoCirculante - inventarios) / pasivoCirculante : 0;

            // Mostrar resultados en Labels
            lblCapitalDeTrabajo.Text = $"Capital de Trabajo: {capitalDeTrabajo:C}";
            lblRazonCirculante.Text = $"Razón Circulante: {razonCirculante:F2}";
            lblRazonRapida.Text = $"Razón Rápida: {razonRapida:F2}";
        }


        private List<DatosEstadoResultado> estadoResultados()
        {
            return new List<DatosEstadoResultado>
            {
                new DatosEstadoResultado { Descripcion = "Ventas Netas", Monto = "$10,198" },
                new DatosEstadoResultado { Descripcion = "Otros Ingresos de Opeacion", Monto = "$113" },
                new DatosEstadoResultado { Descripcion = "Ingresos Totales", Monto = "$10,311" },
                new DatosEstadoResultado { Descripcion = "Costos Totales", Monto = "$5,671" },
                new DatosEstadoResultado { Descripcion = "Utilidad Bruta", Monto = "$4,640" },
                new DatosEstadoResultado { Descripcion = "Gastos de Administracion", Monto = "$447" },
                new DatosEstadoResultado { Descripcion = "Gastos de Venta", Monto = "$2,764" },
                new DatosEstadoResultado { Descripcion = "Otros Ingresos", Monto = "$100" },
                new DatosEstadoResultado { Descripcion = "Otros Gastos", Monto = "232" },
                new DatosEstadoResultado { Descripcion = "Gastos por Interes", Monto = "$366" },
                new DatosEstadoResultado { Descripcion = "Ingresos por Interes", Monto = "$65" },
                new DatosEstadoResultado { Descripcion = "(Perdida)Ganancia por fluctacion cambiaria,neta", Monto = "$(17)" },
                new DatosEstadoResultado { Descripcion = "Utilidad por composicion monetaria por subsidiarias en economia hiperinflacionarias", Monto = "$12" },
                new DatosEstadoResultado { Descripcion = "(Perdida)Ganancia en valor de mercado de instrumentos financieros", Monto = "$(15)" },
                new DatosEstadoResultado { Descripcion = "Utilidad(Perdida) antes de impuesto de operaciones continuas y participacion en los resultados de asociadas y negocios conjuntos", Monto = "$976" },
                new DatosEstadoResultado { Descripcion = "Impuesto a la Utilidad", Monto = "$399" },
                new DatosEstadoResultado { Descripcion = "Participacion en los resultados de Inversiones de capital, neta de impuestos", Monto = "$(7)" },
                new DatosEstadoResultado { Descripcion = "Utilidad(Perdida)neta de operaciones continuas", Monto = "$670" },
                new DatosEstadoResultado { Descripcion = "Utilidad(Perdida) neta consolidada", Monto = "$670" },
                new DatosEstadoResultado { Descripcion = "Valuacion de la porcion efectiva de instrumentos financieros derivados,neto de impuestos", Monto = "$(44)" },
                new DatosEstadoResultado { Descripcion = " Diferencias por efecto de conversion de operaciones y asociaas extranjeras", Monto = "$(296)" },
                new DatosEstadoResultado { Descripcion = "Otras partidas integrales netas que no sean reclasificadas a la (perdida) utilidad en periodos subsecuentes", Monto = "$(340)" },
                new DatosEstadoResultado { Descripcion = "Perdidas de activos financieros de patrimonio clasificado en FVOCI", Monto = "$(11)" },
                new DatosEstadoResultado { Descripcion = "Remedicion del pasivo neto por beneficios definidos neto de impuestos", Monto = "$(28)" },
                new DatosEstadoResultado { Descripcion = "Utilidad en periodos subsecuentes", Monto = "$(39)" },
                new DatosEstadoResultado { Descripcion = "Total de otra (perdida) utilidad integral, neta de impuestos", Monto = "(379)" },
                new DatosEstadoResultado { Descripcion = "Utilidad neta de impuestos consolidada por el año", Monto = "$291" }

            };

        }

        private void CalcularRazonesActividad()
        {
            // Variables necesarias
            decimal costoVentas = 0;
            decimal inventarios = 0;
            decimal cuentasPorCobrar = 0;
            decimal ventasNetas = 0;
            decimal activosFijos = 0;
            decimal activosTotales = 0;
            decimal pasivosTotales = 0;
            decimal pasivos_capital_Totales = 0;
            decimal pasivosNoCorrientesTotales = 0;
            decimal capitalSocial = 0;
            decimal utilidadOperativa = 0;
            decimal gastosIntereses = 0;
            decimal cuentaLarga = 0;
            decimal utilidadBruta = 0;
            decimal utilidadFinal = 0;

            // Obtener valores del estado de resultados
            foreach (var item in DatosFinancieros.EstadoResultados)
            {
                switch (item.Descripcion)
                {
                    case "Costos Totales":
                        costoVentas = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                    case "Ventas Netas":
                        ventasNetas = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                    case "Utilidad(Perdida) neta consolidada":
                        utilidadOperativa = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                    case "Gastos por Interes":
                        gastosIntereses = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                    case "Utilidad Bruta":
                        utilidadBruta = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                    case "Utilidad(Perdida) antes de impuesto de operaciones continuas y participacion en los resultados de asociadas y negocios conjuntos":
                        cuentaLarga = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                    case "Utilidad neta de impuestos consolidada por el año":
                        utilidadFinal = Convert.ToDecimal(item.Monto.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        break;
                }
            }

            // Obtener valores del balance general
            foreach (var item in DatosFinancieros.BalanceGeneral)
            {
                switch (item.Descripcion)
                {
                    case "Inventarios":
                        inventarios = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Cuentas por cobrar, neto":
                        cuentasPorCobrar = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Propiedad, planta y equipo, neto":
                        activosFijos = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "TOTAL ACTIVOS":
                        activosTotales = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Total pasivos":
                        pasivosTotales = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "TOTAL PASIVOS Y CAPITAL":
                        pasivos_capital_Totales = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Total pasivos no corrientes":
                        pasivosNoCorrientesTotales = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;
                    case "Capital Social":
                        capitalSocial = Convert.ToDecimal(item.Monto.Replace(",", ""));
                        break;

                }
            }



            // Cálculos
            decimal rotacionInventarios = inventarios != 0 ? costoVentas / inventarios : 0;
            decimal periodoPromedioCobro = cuentasPorCobrar != 0 && ventasNetas != 0 ? cuentasPorCobrar / (ventasNetas / 360) : 0;
            decimal rotacionCuentasPorCobrar = cuentasPorCobrar != 0 ? ventasNetas / cuentasPorCobrar : 0;
            decimal rotacionActivosFijos = activosFijos != 0 ? ventasNetas / activosFijos : 0;
            decimal rotacionActivosTotales = activosTotales != 0 ? ventasNetas / activosTotales : 0;
            decimal razonDeudaTotal = pasivos_capital_Totales != 0 ? pasivosTotales / pasivos_capital_Totales : 0;
            decimal razonPasivoCapital = capitalSocial != 0 ? pasivosNoCorrientesTotales / capitalSocial : 0;
            decimal rotacionInteresUtilidad = gastosIntereses != 0 ? utilidadOperativa / gastosIntereses : 0;
            decimal MUB = ventasNetas != 0 ? utilidadBruta / ventasNetas : 0;
            decimal MUO = ventasNetas != 0 ? cuentaLarga / ventasNetas : 0;
            decimal MUN = ventasNetas != 0 ? utilidadFinal / ventasNetas : 0;

            // Mostrar resultados en Labels
            lblRotacionInventarios.Text = $"Rotación de Inventarios: {rotacionInventarios:F2}";
            lblPeriodoPromedioCobro.Text = $"Período Promedio de Cobro: {periodoPromedioCobro:F2} días";
            lblRotacionCuentasPorCobrar.Text = $"Rotación de Cuentas por Cobrar: {rotacionCuentasPorCobrar:F2}";
            lblRotacionActivosFijos.Text = $"Rotación de Activos Fijos: {rotacionActivosFijos:F2}";
            lblRotacionActivosTotales.Text = $"Rotación de Activos Totales: {rotacionActivosTotales:F2}";
            lblDeudaTotal.Text = $"Razon de Deuda Total: {razonDeudaTotal:F2}";
            lblPasivoCapital.Text = $"Razon Pasivo a Capital: {razonPasivoCapital:F2}";
            lblInteresUtilidad.Text = $"Razon interes a utilidad: {rotacionInteresUtilidad:F2}";
            lblMUB.Text = $"MUB: {MUB:F2}";
            lblMUO.Text = $"MUO: {MUO:F2}";
            lblMUN.Text = $"MUN: {MUN:F2}";
        }

        public void GenerarBalance(object sender, EventArgs e)
        {
            var balance = BalanceGeneral();
            dgvBalanceGeneral.Rows.Clear();
            foreach (var item in balance)
            {
                dgvBalanceGeneral.Rows.Add(item.Descripcion, item.Monto);
            }

            DatosFinancieros.BalanceGeneral = BalanceGeneral();
            ActualizarGrids();
        }

        public void GenerarEstado(object sender, EventArgs e)
        {
            var estado = estadoResultados();
            dgvEstadoDeResultado.Rows.Clear();
            foreach (var item in estado)
            {
                dgvEstadoDeResultado.Rows.Add(item.Descripcion, item.Monto);
            }

            DatosFinancieros.EstadoResultados = estadoResultados();
            ActualizarGrids();
        }

        public Dictionary<string, int> ObtenerDatos()
        {
            var datos = new Dictionary<string, int>();
            try
            {
                // Procesar datos del dgvBalanceGeneral
                foreach (DataGridViewRow row in dgvBalanceGeneral.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                    {
                        string nombreCuenta = row.Cells[0].Value.ToString();
                        int valorCuenta = Convert.ToInt32(row.Cells[1].Value);
                        datos[nombreCuenta] = valorCuenta;
                    }
                }

                // Procesar datos del dgvEstadoResultados
                foreach (DataGridViewRow row in dgvEstadoDeResultado.Rows)
                {
                    if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                    {
                        string nombreCuenta = row.Cells[0].Value.ToString();
                        int valorCuenta = Convert.ToInt32(row.Cells[1].Value);

                        // Si ya existe una clave con el mismo nombre, sumar los valores
                        if (datos.ContainsKey(nombreCuenta))
                        {
                            datos[nombreCuenta] += valorCuenta;
                        }
                        else
                        {
                            datos[nombreCuenta] = valorCuenta;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener datos: {ex.Message}");
            }
            return datos;
        }

        private void FrmBalanceYEstado_Activated(object sender, EventArgs e)
        {
            ActualizarGrids();
        }

        // Variable global para almacenar la aplicación Excel y el libro
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook;


        private void btnExportarBalance_Click(object sender, EventArgs e)
        {
            // ESTE FUE UNA PRUEBA EXITOSA DE EXCEL, DEJAR COMENTADO POR SI ACASO 

            //// Crear un nuevo libro de Excel
            //using (var workbook = new XLWorkbook())
            //{
            //    // Crear una nueva hoja para el Balance General
            //    var worksheet = workbook.Worksheets.Add("Balance General");

            //    // Exportar los encabezados del DataGridView
            //    for (int i = 0; i < dgvBalanceGeneral.Columns.Count; i++)
            //    {
            //        worksheet.Cell(1, i + 1).Value = dgvBalanceGeneral.Columns[i].HeaderText;
            //    }

            //    // Exportar las filas del DataGridView
            //    for (int i = 0; i < dgvBalanceGeneral.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < dgvBalanceGeneral.Columns.Count; j++)
            //        {
            //            // Obtener el valor de la celda y convertirlo explícitamente
            //            object cellValue = dgvBalanceGeneral.Rows[i].Cells[j].Value;
            //            worksheet.Cell(i + 2, j + 1).Value = cellValue != null ? cellValue.ToString() : string.Empty;
            //        }
            //    }

            //    // Guardar el archivo
            //    using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" })
            //    {
            //        if (sfd.ShowDialog() == DialogResult.OK)
            //        {
            //            workbook.SaveAs(sfd.FileName);
            //            MessageBox.Show("Exportación completada con éxito.", "Exportar a Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        }
            //    }
            //}


            //OTRO CORRECTO

            //// Crear una aplicación Excel
            //Excel.Application excelApp = new Excel.Application();

            //try
            //{
            //    // Crear un nuevo libro de Excel
            //    Excel.Workbook workbook = excelApp.Workbooks.Add();
            //    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            //    // Títulos de las columnas
            //    for (int i = 0; i < dgvBalanceGeneral.Columns.Count; i++)
            //    {
            //        worksheet.Cells[1, i + 1] = dgvBalanceGeneral.Columns[i].HeaderText;
            //    }

            //    // Datos de las filas
            //    for (int i = 0; i < dgvBalanceGeneral.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < dgvBalanceGeneral.Columns.Count; j++)
            //        {
            //            object cellValue = dgvBalanceGeneral.Rows[i].Cells[j].Value;
            //            worksheet.Cells[i + 2, j + 1] = cellValue != null ? cellValue.ToString() : string.Empty;
            //        }
            //    }

            //    // Ajustar columnas automáticamente
            //    worksheet.Columns.AutoFit();

            //    // Mostrar Excel al usuario
            //    excelApp.Visible = true;

            //    // Si quieres guardar el archivo automáticamente, descomenta las siguientes líneas:
            //    /*
            //    string filePath = "C:\\Users\\Public\\BalanceGeneral.xlsx"; // Cambia esto según tus necesidades
            //    workbook.SaveAs(filePath);
            //    MessageBox.Show("Archivo guardado en: " + filePath, "Exportación completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    */
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    // Libera los recursos de COM (no cierres si necesitas Excel abierto)
            //    if (excelApp != null)
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //}

            // Si Excel no está abierto, abrirlo
            if (excelApp == null)
                excelApp = new Excel.Application();

            try
            {
                // Si el libro no está creado, lo creamos
                if (workbook == null)
                    workbook = excelApp.Workbooks.Add();

                // Crear una hoja para Balance General
                Excel.Worksheet worksheetBalance = (Excel.Worksheet)workbook.Sheets.Add();
                worksheetBalance.Name = "Balance General"; // Asignamos un nombre a la hoja

                // Títulos de las columnas del Balance General
                for (int i = 0; i < dgvBalanceGeneral.Columns.Count; i++)
                {
                    worksheetBalance.Cells[1, i + 1] = dgvBalanceGeneral.Columns[i].HeaderText;
                }

                // Datos de las filas del Balance General
                for (int i = 0; i < dgvBalanceGeneral.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvBalanceGeneral.Columns.Count; j++)
                    {
                        object cellValue = dgvBalanceGeneral.Rows[i].Cells[j].Value;
                        worksheetBalance.Cells[i + 2, j + 1] = cellValue != null ? cellValue.ToString() : string.Empty;
                    }
                }

                // Ajustar columnas automáticamente
                worksheetBalance.Columns.AutoFit();

                // Mostrar Excel al usuario
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportarER_Click(object sender, EventArgs e)
        {
            //// Crear una aplicación Excel
            //Excel.Application excelApp = new Excel.Application();

            //try
            //{
            //    // Crear un nuevo libro de Excel
            //    Excel.Workbook workbook = excelApp.Workbooks.Add();
            //    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            //    // Títulos de las columnas
            //    worksheet.Cells[1, 1] = "Descripción";
            //    worksheet.Cells[1, 2] = "Monto";

            //    // Datos de las filas (usando la lista de estado de resultado)
            //    List<DatosEstadoResultado> estadoResultadoList = estadoResultados();
            //    for (int i = 0; i < estadoResultadoList.Count; i++)
            //    {
            //        worksheet.Cells[i + 2, 1] = estadoResultadoList[i].Descripcion;
            //        worksheet.Cells[i + 2, 2] = estadoResultadoList[i].Monto;
            //    }

            //    // Ajustar columnas automáticamente
            //    worksheet.Columns.AutoFit();

            //    // Mostrar Excel al usuario
            //    excelApp.Visible = true;

            //    // Si quieres guardar el archivo automáticamente, descomenta las siguientes líneas:
            //    /*
            //    string filePath = "C:\\Users\\Public\\EstadoResultado.xlsx"; // Cambia esto según tus necesidades
            //    workbook.SaveAs(filePath);
            //    MessageBox.Show("Archivo guardado en: " + filePath, "Exportación completada", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    */
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
            //finally
            //{
            //    // Libera los recursos de COM (no cierres si necesitas Excel abierto)
            //    if (excelApp != null)
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //}

            // Si Excel no está abierto, abrirlo
            if (excelApp == null)
                excelApp = new Excel.Application();

            try
            {
                // Si el libro no está creado, lo creamos
                if (workbook == null)
                    workbook = excelApp.Workbooks.Add();

                // Crear una hoja para Estado de Resultado
                Excel.Worksheet worksheetEstado = (Excel.Worksheet)workbook.Sheets.Add();
                worksheetEstado.Name = "Estado de Resultado"; // Asignamos un nombre a la hoja

                // Títulos de las columnas del Estado de Resultado
                worksheetEstado.Cells[1, 1] = "Descripción";
                worksheetEstado.Cells[1, 2] = "Monto";

                // Datos de las filas del Estado de Resultado
                List<DatosEstadoResultado> estadoResultadoList = estadoResultados();
                for (int i = 0; i < estadoResultadoList.Count; i++)
                {
                    worksheetEstado.Cells[i + 2, 1] = estadoResultadoList[i].Descripcion;
                    worksheetEstado.Cells[i + 2, 2] = estadoResultadoList[i].Monto;
                }

                // Ajustar columnas automáticamente
                worksheetEstado.Columns.AutoFit();

                // Mostrar Excel al usuario
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Función para exportar Balance General
        private void ExportarAExcelBalanceGeneral()
        {
            // Crear una aplicación Excel
            Excel.Application excelApp = new Excel.Application();

            try
            {
                // Crear un nuevo libro de Excel
                Excel.Workbook workbook = excelApp.Workbooks.Add();

                // Crear una hoja para Balance General
                Excel.Worksheet worksheetBalance = (Excel.Worksheet)workbook.Sheets.Add();
                worksheetBalance.Name = "Balance General"; // Asignamos un nombre a la hoja

                // Títulos de las columnas del Balance General
                for (int i = 0; i < dgvBalanceGeneral.Columns.Count; i++)
                {
                    worksheetBalance.Cells[1, i + 1] = dgvBalanceGeneral.Columns[i].HeaderText;
                }

                // Datos de las filas del Balance General
                for (int i = 0; i < dgvBalanceGeneral.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvBalanceGeneral.Columns.Count; j++)
                    {
                        object cellValue = dgvBalanceGeneral.Rows[i].Cells[j].Value;
                        worksheetBalance.Cells[i + 2, j + 1] = cellValue != null ? cellValue.ToString() : string.Empty;
                    }
                }

                // Ajustar columnas automáticamente
                worksheetBalance.Columns.AutoFit();

                // Mostrar Excel al usuario
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Libera los recursos de COM (no cierres si necesitas Excel abierto)
                if (excelApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        // Función para exportar Estado de Resultado
        private void ExportarAExcelEstadoResultado()
        {
            // Crear una aplicación Excel
            Excel.Application excelApp = new Excel.Application();

            try
            {
                // Crear un nuevo libro de Excel
                Excel.Workbook workbook = excelApp.Workbooks.Add();

                // Crear una hoja para Estado de Resultado
                Excel.Worksheet worksheetEstado = (Excel.Worksheet)workbook.Sheets.Add();
                worksheetEstado.Name = "Estado de Resultado"; // Asignamos un nombre a la hoja

                // Títulos de las columnas del Estado de Resultado
                worksheetEstado.Cells[1, 1] = "Descripción";
                worksheetEstado.Cells[1, 2] = "Monto";

                // Datos de las filas del Estado de Resultado
                List<DatosEstadoResultado> estadoResultadoList = estadoResultados();
                for (int i = 0; i < estadoResultadoList.Count; i++)
                {
                    worksheetEstado.Cells[i + 2, 1] = estadoResultadoList[i].Descripcion;
                    worksheetEstado.Cells[i + 2, 2] = estadoResultadoList[i].Monto;
                }

                // Ajustar columnas automáticamente
                worksheetEstado.Columns.AutoFit();

                // Mostrar Excel al usuario
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Libera los recursos de COM (no cierres si necesitas Excel abierto)
                if (excelApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
    