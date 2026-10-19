namespace ComplementaCadastro
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btn_Importar = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.btnConectar = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnAtualizaCPF = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.btn_InsereEnvolvido = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.txtMaquina = new System.Windows.Forms.ToolStripTextBox();
            this.btnDistribuir = new System.Windows.Forms.ToolStripButton();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.StatusLabel_Registros = new System.Windows.Forms.ToolStripStatusLabel();
            this.grade = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.grid_dados = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_InserePjReu = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grid_dados)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Font = new System.Drawing.Font("Bradesco Sans", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btn_Importar,
            this.toolStripSeparator6,
            this.btnConectar,
            this.toolStripSeparator1,
            this.btnAtualizaCPF,
            this.toolStripSeparator3,
            this.btn_InsereEnvolvido,
            this.toolStripSeparator4,
            this.btn_InserePjReu,
            this.toolStripSeparator2,
            this.toolStripLabel1,
            this.txtMaquina,
            this.btnDistribuir});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1223, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // btn_Importar
            // 
            this.btn_Importar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btn_Importar.Image = ((System.Drawing.Image)(resources.GetObject("btn_Importar.Image")));
            this.btn_Importar.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btn_Importar.Name = "btn_Importar";
            this.btn_Importar.Size = new System.Drawing.Size(107, 22);
            this.btn_Importar.Text = "Importar planilha";
            this.btn_Importar.Click += new System.EventHandler(this.btn_Importar_Click);
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(6, 25);
            // 
            // btnConectar
            // 
            this.btnConectar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnConectar.Image = ((System.Drawing.Image)(resources.GetObject("btnConectar.Image")));
            this.btnConectar.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnConectar.Name = "btnConectar";
            this.btnConectar.Size = new System.Drawing.Size(92, 22);
            this.btnConectar.Text = "Conectar GCPJ";
            this.btnConectar.Click += new System.EventHandler(this.btnConectar_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // btnAtualizaCPF
            // 
            this.btnAtualizaCPF.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnAtualizaCPF.Image = ((System.Drawing.Image)(resources.GetObject("btnAtualizaCPF.Image")));
            this.btnAtualizaCPF.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnAtualizaCPF.Name = "btnAtualizaCPF";
            this.btnAtualizaCPF.Size = new System.Drawing.Size(312, 22);
            this.btnAtualizaCPF.Text = "Atualizar CPF/CNPJ Envolvido (GCPJ Ativo e Encerrado)";
            this.btnAtualizaCPF.Click += new System.EventHandler(this.btnAtualizaCPF_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
            // 
            // btn_InsereEnvolvido
            // 
            this.btn_InsereEnvolvido.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btn_InsereEnvolvido.Image = ((System.Drawing.Image)(resources.GetObject("btn_InsereEnvolvido.Image")));
            this.btn_InsereEnvolvido.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btn_InsereEnvolvido.Name = "btn_InsereEnvolvido";
            this.btn_InsereEnvolvido.Size = new System.Drawing.Size(208, 22);
            this.btn_InsereEnvolvido.Text = "Inserir Demais Autores (GCPJ Ativo)";
            this.btn_InsereEnvolvido.Click += new System.EventHandler(this.btn_InsereEnvolvido_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(53, 22);
            this.toolStripLabel1.Text = "Máquina";
            // 
            // txtMaquina
            // 
            this.txtMaquina.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtMaquina.Name = "txtMaquina";
            this.txtMaquina.Size = new System.Drawing.Size(28, 25);
            this.txtMaquina.Text = "1";
            this.txtMaquina.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // btnDistribuir
            // 
            this.btnDistribuir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnDistribuir.Image = ((System.Drawing.Image)(resources.GetObject("btnDistribuir.Image")));
            this.btnDistribuir.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnDistribuir.Name = "btnDistribuir";
            this.btnDistribuir.Size = new System.Drawing.Size(118, 22);
            this.btnDistribuir.Text = "Distribuir Máquinas";
            this.btnDistribuir.Click += new System.EventHandler(this.btnDistribuir_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.StatusLabel_Registros});
            this.statusStrip1.Location = new System.Drawing.Point(0, 737);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1223, 24);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.BorderSides = ((System.Windows.Forms.ToolStripStatusLabelBorderSides)((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left | System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom)));
            this.toolStripStatusLabel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(1202, 19);
            this.toolStripStatusLabel1.Spring = true;
            this.toolStripStatusLabel1.Text = "BANCO BRADESCO S/A - DEPARTAMENTO JURÍDICO - MATRIZ";
            this.toolStripStatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // StatusLabel_Registros
            // 
            this.StatusLabel_Registros.BorderSides = ((System.Windows.Forms.ToolStripStatusLabelBorderSides)((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left | System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom)));
            this.StatusLabel_Registros.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.StatusLabel_Registros.Name = "StatusLabel_Registros";
            this.StatusLabel_Registros.Size = new System.Drawing.Size(4, 19);
            // 
            // grade
            // 
            this.grade.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grade.Font = new System.Drawing.Font("Bradesco Sans", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grade.FormattingEnabled = true;
            this.grade.ItemHeight = 15;
            this.grade.Location = new System.Drawing.Point(13, 25);
            this.grade.Name = "grade";
            this.grade.Size = new System.Drawing.Size(1168, 349);
            this.grade.TabIndex = 2;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.linkLabel1);
            this.groupBox1.Controls.Add(this.grid_dados);
            this.groupBox1.Location = new System.Drawing.Point(14, 447);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1195, 275);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Registros importados";
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.LinkColor = System.Drawing.Color.Green;
            this.linkLabel1.Location = new System.Drawing.Point(1068, 9);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(114, 15);
            this.linkLabel1.TabIndex = 1;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Exportar para Excel";
            this.linkLabel1.VisitedLinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // grid_dados
            // 
            this.grid_dados.AllowUserToAddRows = false;
            this.grid_dados.AllowUserToDeleteRows = false;
            this.grid_dados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grid_dados.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.grid_dados.BackgroundColor = System.Drawing.Color.White;
            this.grid_dados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.grid_dados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid_dados.Location = new System.Drawing.Point(13, 25);
            this.grid_dados.Name = "grid_dados";
            this.grid_dados.ReadOnly = true;
            this.grid_dados.RowHeadersWidth = 21;
            this.grid_dados.Size = new System.Drawing.Size(1168, 234);
            this.grid_dados.TabIndex = 0;
            this.grid_dados.CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.grid_dados_CellMouseUp);
            this.grid_dados.KeyUp += new System.Windows.Forms.KeyEventHandler(this.grid_dados_KeyUp);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.grade);
            this.groupBox2.Font = new System.Drawing.Font("Bradesco Sans", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(14, 32);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1195, 391);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Status do processamento";
            // 
            // btn_InserePjReu
            // 
            this.btn_InserePjReu.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btn_InserePjReu.Image = ((System.Drawing.Image)(resources.GetObject("btn_InserePjReu.Image")));
            this.btn_InserePjReu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btn_InserePjReu.Name = "btn_InserePjReu";
            this.btn_InserePjReu.Size = new System.Drawing.Size(86, 22);
            this.btn_InserePjReu.Text = "Inserir PJ RÉU";
            this.btn_InserePjReu.Click += new System.EventHandler(this.btn_InserePjReu_Click);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 25);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1223, 761);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.toolStrip1);
            this.Font = new System.Drawing.Font("Bradesco Sans", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Form1";
            this.Text = "ATUALIZA ENVOLVIDOS";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grid_dados)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripButton btnConectar;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnAtualizaCPF;
        private System.Windows.Forms.ListBox grade;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripTextBox txtMaquina;
        private System.Windows.Forms.ToolStripButton btnDistribuir;
        private System.Windows.Forms.ToolStripButton btn_Importar;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView grid_dados;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel_Registros;
        private System.Windows.Forms.ToolStripButton btn_InsereEnvolvido;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton btn_InserePjReu;
    }
}

