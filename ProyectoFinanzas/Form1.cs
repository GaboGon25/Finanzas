using System.Runtime.InteropServices;

namespace ProyectoFinanzas
{
    public partial class Form1 : Form
    {
        private FrmBalanceYEstado frmBalanceYEstado;

        public Form1()
        {
            InitializeComponent();
        }

        private void Cerrar(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Maximizar(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            btnMaximizar.Visible = false;
            btnRestaurar.Visible = true;
        }

        private void Restaurar(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            btnRestaurar.Visible = false;
            btnMaximizar.Visible = true;
        }

        private void Minimizar(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        [DllImport("user32.dll", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void BarraTitulo_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void AbrirFormHija(object formHija)
        {
            if (this.panelContenedor.Controls.Count > 0)
                this.panelContenedor.Controls.RemoveAt(0);
            Form fh = formHija as Form;
            fh.TopLevel = false;
            fh.Dock = DockStyle.Fill;
            this.panelContenedor.Controls.Add(fh);
            this.panelContenedor.Tag = fh;
            fh.Show();
        }

        private void BalanceYEstadoResultado(object sender, EventArgs e)
        {
            AbrirFormHija(new FrmBalanceYEstado());
        }

        private void Presentancion(object sender, EventArgs e)
        {
            AbrirFormHija(new FrmInicio());
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Presentancion(null, e);
        }

        private void btnRazones_Click(object sender, EventArgs e)
        {
            AbrirFormHija(new frmRazones(frmBalanceYEstado));

        }
    }
}
