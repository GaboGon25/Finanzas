using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;

namespace ProyectoFinanzas
{
    public partial class frmRazones : Form
    {
        private FrmBalanceYEstado frmBalanceYEstado; // Referencia al Form1
        private Form1 form1;

        //Datos para las razones de liquidez
        private decimal ActivoCorriente = 3011;
        private decimal PasivoCorriente = 2705;
        private decimal Inventarios = 559;
        private decimal CapitalTrabajoNeto;
        private decimal RazonCirculante;
        private decimal RazonRapida;

        //Datos para las razones de actividad
        private decimal CostoArticulosVendidos;
        private decimal RotacionInventarios;
        private decimal RotacionCuentasXCobrar;
        private decimal VentaCredito;
        private decimal RazonActivosFijos;
        private decimal RazonActivosTotales;
        private decimal ActivosTotales = 13671;
        private decimal CuentasXCobrar = 821;
        private decimal PeriodoPromedioCobro;
        private decimal Ventas = 10198;

        //Datos para las razones de endeudamiento
        private decimal PasivosTotales;
        private decimal PasivoL_P;
        private decimal CapitalSocial;
        private decimal UtilidadEInteresAntesImpuesto;
        private decimal CargosXIntereses;
        private decimal RazonDeudaTotal;
        private decimal RazonPasivoCapital;
        private decimal RazonInteres_Utilidad;

        //Datos para las razones de rentabilidad
        private decimal MUB;
        private decimal MUO;
        private decimal MUN;
        private decimal UtilidadBruta;
        private decimal UtilidadOperativa;
        private decimal UtilidadNeta;






        public frmRazones(FrmBalanceYEstado frmBalanceYEstado)
        {
            InitializeComponent();
            this.frmBalanceYEstado = frmBalanceYEstado;
        }

        public frmRazones(Form1 form1)
        {
            this.form1 = form1;
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void btnLiquidez_Click(object sender, EventArgs e)
        {

            // Calculamos el Capital de Trabajo Neto
            CapitalTrabajoNeto = ActivoCorriente - PasivoCorriente;

            // Calculamos la Razón Circulante
            RazonCirculante = ActivoCorriente / PasivoCorriente;

            // Calculamos la Razón Rápida
            RazonRapida = (ActivoCorriente - Inventarios) / PasivoCorriente;

            // resultados en los Labels
            lblCapitalNeto.Text = "Capital de Trabajo Neto: " + CapitalTrabajoNeto.ToString("F2");
            lblRazonCirculante.Text = "Razón Circulante: " + RazonCirculante.ToString("F2");
            lblRazonRapida.Text = "Razón Rápida: " + RazonRapida.ToString("F2");
        }

        private void btnActividad_Click(object sender, EventArgs e)
        {
            CostoArticulosVendidos = 5671;
            VentaCredito = 10311;
            //Calculamos la razon de rotacion de inventarios
            RotacionInventarios = CostoArticulosVendidos / Inventarios;

            //Calculamos la razon de rotacion de cuentas por cobrar
            RotacionCuentasXCobrar = VentaCredito / CuentasXCobrar;

            //Calculamos la razon de periodo promedio de cobro
            PeriodoPromedioCobro = CuentasXCobrar / (Ventas / 360);

            //Calculamos la razon de rotacion de activos fijos
            RazonActivosFijos = Ventas / ActivoCorriente;

            //Calculamos la razon de rotacion de activos totales
            RazonActivosTotales = Ventas / ActivosTotales;


            // resultados en los Labels
            lblInventario.Text = "Rotacion de Inventarios: " + RotacionInventarios.ToString("F2");
            lblCuentasXCobrar.Text = "Rotacion de cuentas por cobrar: " + RotacionCuentasXCobrar.ToString("F2");
            lblCobro.Text = "Periodo Promedio de cobro (dias): " + PeriodoPromedioCobro.ToString("F2");
            lblActivosFijos.Text = "Rotacion de activos fijos: " + RazonActivosFijos.ToString("F2");
            lblActivosTotales.Text = "Rotacion de activos totales: " + RazonActivosTotales.ToString("F2");
        }

        private void btnEndeudamiento_Click(object sender, EventArgs e)
        {
            PasivosTotales = 6795;
            PasivoL_P = 4090;
            CapitalSocial = 109;
            UtilidadEInteresAntesImpuesto = 670;
            CargosXIntereses = 366;
            //Calculamos la deuda total
            RazonDeudaTotal = PasivosTotales / ActivosTotales;

            //Calculamos pasivo a capital
            RazonPasivoCapital = PasivoL_P / CapitalSocial;

            //Calculamos interes a utilidad
            RazonInteres_Utilidad = UtilidadEInteresAntesImpuesto / CargosXIntereses;

            // resultados en los Labels
            lblDeudaTotal.Text = "Deuda Total: " + RazonDeudaTotal.ToString("F2");
            lblPasivoCapital.Text = "Pasivo a Capital: " + RazonPasivoCapital.ToString("F2");
            lblIntereses.Text = "Rotacion de interes a utilidades: " + RazonInteres_Utilidad.ToString("F2");
        }

        private void btnRentabilidad_Click(object sender, EventArgs e)
        {
            UtilidadBruta = 4640;
            UtilidadOperativa = 976;
            UtilidadNeta = 291;

            //Calculo del MUB
            MUB = UtilidadBruta / Ventas;

            //Calculo del MUO
            MUO = UtilidadOperativa / Ventas;

            //Calculo del MUN
            MUN = UtilidadNeta / Ventas;

            lblMUB.Text = "MUB: " + MUB.ToString("F2");
            lblMUO.Text = "MUO: " + MUO.ToString("F2");
            lblMUN.Text = "MUN: " + MUN.ToString("F2");
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
    }



}
