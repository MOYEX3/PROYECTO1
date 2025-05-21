namespace proyecto1
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            txtBusqueda = new TextBox();
            btBuscar = new Button();
            richTextBox1 = new RichTextBox();
            label1 = new Label();
            label2 = new Label();
            SuspendLayout();
            // 
            // txtBusqueda
            // 
            txtBusqueda.Location = new Point(39, 200);
            txtBusqueda.Name = "txtBusqueda";
            txtBusqueda.Size = new Size(345, 27);
            txtBusqueda.TabIndex = 0;
            // 
            // btBuscar
            // 
            btBuscar.Location = new Point(115, 278);
            btBuscar.Name = "btBuscar";
            btBuscar.Size = new Size(183, 65);
            btBuscar.TabIndex = 1;
            btBuscar.Text = "Buscar";
            btBuscar.UseVisualStyleBackColor = true;
            btBuscar.Click += btBuscar_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(433, 80);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(494, 458);
            richTextBox1.TabIndex = 2;
            richTextBox1.Text = "";
            richTextBox1.TextChanged += richTextBox1_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Showcard Gothic", 16.2F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.Location = new Point(12, 108);
            label1.Name = "label1";
            label1.Size = new Size(377, 35);
            label1.TabIndex = 3;
            label1.Text = "haga su consulta aqui ";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Showcard Gothic", 16.2F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label2.Location = new Point(494, 24);
            label2.Name = "label2";
            label2.Size = new Size(398, 35);
            label2.TabIndex = 3;
            label2.Text = "resultado de busqueda ";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            BackgroundImageLayout = ImageLayout.Zoom;
            ClientSize = new Size(939, 550);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(richTextBox1);
            Controls.Add(btBuscar);
            Controls.Add(txtBusqueda);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtBusqueda;
        private Button btBuscar;
        private RichTextBox richTextBox1;
        private Label label1;
        private Label label2;
    }
}
