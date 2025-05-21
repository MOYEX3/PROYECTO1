namespace proyecto1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private async void btBuscar_Click(object sender, EventArgs e)
        {
            var class1 = new Class1();
            string apikey = "colocar aqui ";
            string promt = txtBusqueda.Text;

            string respuesta = await class1.BuscarenopeniaAsync(promt, apikey);
            richTextBox1.Text = respuesta;

            string ruta = class1.crearcarpetain();
            class1.documentoword(ruta, "Investigación: " + promt, respuesta);
            class1.documentopptx(ruta, "Resumen: " + promt, respuesta);
            class1.GuardarEnBaseDeDatos(promt, respuesta);
            MessageBox.Show("¡Documentos generados exitosamente en:\n" + ruta, "Éxito");
        }
    }
}
