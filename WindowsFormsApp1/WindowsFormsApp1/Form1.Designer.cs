namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Variavel de designer necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estao sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessario descartar os recursos gerenciados; caso contrario, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codigo gerado pelo Windows Form Designer

        /// <summary>
        /// Metodo necessario para suporte ao Designer - nao modifique
        /// o conteudo deste metodo com o editor de codigo.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.pnlTeklaIndicator = new System.Windows.Forms.Panel();
            this.lblTeklaStatusLink = new System.Windows.Forms.Label();
            this.pnlDrawingIndicator = new System.Windows.Forms.Panel();
            this.lblDrawingStatusLink = new System.Windows.Forms.Label();
            this.chkPlanoZY = new System.Windows.Forms.CheckBox();
            this.chkPlanoXZ = new System.Windows.Forms.CheckBox();
            this.chkPlanoXY = new System.Windows.Forms.CheckBox();
            this.chkGhostLinhas = new System.Windows.Forms.CheckBox();
            this.chkLinhas = new System.Windows.Forms.CheckBox();
            this.chkColorir = new System.Windows.Forms.CheckBox();
            this.btnCriarFolhaInteira = new System.Windows.Forms.Button();
            this.btnCriarAreaDefinida = new System.Windows.Forms.Button();
            this.btnLimparVistaExplodida = new System.Windows.Forms.Button();
            this.lblStatusTekla = new System.Windows.Forms.Label();
            this.btnToggleLog = new System.Windows.Forms.Button();
            this.pnlLog = new System.Windows.Forms.Panel();
            this.txtSaida = new System.Windows.Forms.TextBox();
            this.pnlLog.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlTeklaIndicator
            // 
            this.pnlTeklaIndicator.BackColor = System.Drawing.Color.Silver;
            this.pnlTeklaIndicator.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pnlTeklaIndicator.Location = new System.Drawing.Point(22, 20);
            this.pnlTeklaIndicator.Name = "pnlTeklaIndicator";
            this.pnlTeklaIndicator.Size = new System.Drawing.Size(14, 14);
            this.pnlTeklaIndicator.TabIndex = 0;
            this.pnlTeklaIndicator.Click += new System.EventHandler(this.btnVerificarTekla_Click);
            // 
            // lblTeklaStatusLink
            // 
            this.lblTeklaStatusLink.AutoSize = true;
            this.lblTeklaStatusLink.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblTeklaStatusLink.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTeklaStatusLink.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(56)))), ((int)(((byte)(69)))));
            this.lblTeklaStatusLink.Location = new System.Drawing.Point(42, 18);
            this.lblTeklaStatusLink.Name = "lblTeklaStatusLink";
            this.lblTeklaStatusLink.Size = new System.Drawing.Size(49, 16);
            this.lblTeklaStatusLink.TabIndex = 1;
            this.lblTeklaStatusLink.Text = "O Tekla";
            this.lblTeklaStatusLink.Click += new System.EventHandler(this.btnVerificarTekla_Click);
            // 
            // pnlDrawingIndicator
            // 
            this.pnlDrawingIndicator.BackColor = System.Drawing.Color.Silver;
            this.pnlDrawingIndicator.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pnlDrawingIndicator.Location = new System.Drawing.Point(168, 20);
            this.pnlDrawingIndicator.Name = "pnlDrawingIndicator";
            this.pnlDrawingIndicator.Size = new System.Drawing.Size(14, 14);
            this.pnlDrawingIndicator.TabIndex = 2;
            this.pnlDrawingIndicator.Click += new System.EventHandler(this.btnVerificarDesenho_Click);
            // 
            // lblDrawingStatusLink
            // 
            this.lblDrawingStatusLink.AutoSize = true;
            this.lblDrawingStatusLink.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblDrawingStatusLink.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDrawingStatusLink.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(56)))), ((int)(((byte)(69)))));
            this.lblDrawingStatusLink.Location = new System.Drawing.Point(188, 18);
            this.lblDrawingStatusLink.Name = "lblDrawingStatusLink";
            this.lblDrawingStatusLink.Size = new System.Drawing.Size(67, 16);
            this.lblDrawingStatusLink.TabIndex = 3;
            this.lblDrawingStatusLink.Text = "O Desenho";
            this.lblDrawingStatusLink.Click += new System.EventHandler(this.btnVerificarDesenho_Click);
            // 
            // chkPlanoZY
            // 
            this.chkPlanoZY.AutoSize = true;
            this.chkPlanoZY.ForeColor = System.Drawing.Color.Firebrick;
            this.chkPlanoZY.Location = new System.Drawing.Point(22, 58);
            this.chkPlanoZY.Name = "chkPlanoZY";
            this.chkPlanoZY.Size = new System.Drawing.Size(37, 20);
            this.chkPlanoZY.TabIndex = 4;
            this.chkPlanoZY.Text = "X";
            this.chkPlanoZY.UseVisualStyleBackColor = true;
            // 
            // chkPlanoXZ
            // 
            this.chkPlanoXZ.AutoSize = true;
            this.chkPlanoXZ.ForeColor = System.Drawing.Color.ForestGreen;
            this.chkPlanoXZ.Location = new System.Drawing.Point(115, 58);
            this.chkPlanoXZ.Name = "chkPlanoXZ";
            this.chkPlanoXZ.Size = new System.Drawing.Size(36, 20);
            this.chkPlanoXZ.TabIndex = 5;
            this.chkPlanoXZ.Text = "Y";
            this.chkPlanoXZ.UseVisualStyleBackColor = true;
            // 
            // chkPlanoXY
            // 
            this.chkPlanoXY.AutoSize = true;
            this.chkPlanoXY.Checked = true;
            this.chkPlanoXY.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkPlanoXY.ForeColor = System.Drawing.Color.RoyalBlue;
            this.chkPlanoXY.Location = new System.Drawing.Point(208, 58);
            this.chkPlanoXY.Name = "chkPlanoXY";
            this.chkPlanoXY.Size = new System.Drawing.Size(37, 20);
            this.chkPlanoXY.TabIndex = 6;
            this.chkPlanoXY.Text = "Z";
            this.chkPlanoXY.UseVisualStyleBackColor = true;
            // 
            // chkGhostLinhas
            // 
            this.chkGhostLinhas.AutoSize = true;
            this.chkGhostLinhas.Checked = true;
            this.chkGhostLinhas.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkGhostLinhas.Location = new System.Drawing.Point(22, 88);
            this.chkGhostLinhas.Name = "chkGhostLinhas";
            this.chkGhostLinhas.Size = new System.Drawing.Size(64, 20);
            this.chkGhostLinhas.TabIndex = 7;
            this.chkGhostLinhas.Text = "Ghost";
            this.chkGhostLinhas.UseVisualStyleBackColor = true;
            // 
            // chkLinhas
            // 
            this.chkLinhas.AutoSize = true;
            this.chkLinhas.Checked = true;
            this.chkLinhas.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLinhas.Location = new System.Drawing.Point(115, 88);
            this.chkLinhas.Name = "chkLinhas";
            this.chkLinhas.Size = new System.Drawing.Size(67, 20);
            this.chkLinhas.TabIndex = 8;
            this.chkLinhas.Text = "Linhas";
            this.chkLinhas.UseVisualStyleBackColor = true;
            // 
            // chkColorir
            // 
            this.chkColorir.AutoSize = true;
            this.chkColorir.Location = new System.Drawing.Point(208, 88);
            this.chkColorir.Name = "chkColorir";
            this.chkColorir.Size = new System.Drawing.Size(68, 20);
            this.chkColorir.TabIndex = 9;
            this.chkColorir.Text = "Colorir";
            this.chkColorir.UseVisualStyleBackColor = true;
            // 
            // btnCriarFolhaInteira
            // 
            this.btnCriarFolhaInteira.Location = new System.Drawing.Point(22, 126);
            this.btnCriarFolhaInteira.Name = "btnCriarFolhaInteira";
            this.btnCriarFolhaInteira.Size = new System.Drawing.Size(276, 36);
            this.btnCriarFolhaInteira.TabIndex = 10;
            this.btnCriarFolhaInteira.BackColor = System.Drawing.Color.White;
            this.btnCriarFolhaInteira.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCriarFolhaInteira.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCriarFolhaInteira.Text = "Explodir em folha inteira";
            this.btnCriarFolhaInteira.UseVisualStyleBackColor = true;
            this.btnCriarFolhaInteira.Click += new System.EventHandler(this.btnCriarFolhaInteira_Click);
            // 
            // btnCriarAreaDefinida
            // 
            this.btnCriarAreaDefinida.Location = new System.Drawing.Point(22, 170);
            this.btnCriarAreaDefinida.Name = "btnCriarAreaDefinida";
            this.btnCriarAreaDefinida.Size = new System.Drawing.Size(276, 36);
            this.btnCriarAreaDefinida.TabIndex = 11;
            this.btnCriarAreaDefinida.BackColor = System.Drawing.Color.White;
            this.btnCriarAreaDefinida.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCriarAreaDefinida.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCriarAreaDefinida.Text = "Definir area";
            this.btnCriarAreaDefinida.UseVisualStyleBackColor = true;
            this.btnCriarAreaDefinida.Click += new System.EventHandler(this.btnCriarAreaDefinida_Click);
            // 
            // btnLimparVistaExplodida
            // 
            this.btnLimparVistaExplodida.Location = new System.Drawing.Point(22, 214);
            this.btnLimparVistaExplodida.Name = "btnLimparVistaExplodida";
            this.btnLimparVistaExplodida.Size = new System.Drawing.Size(276, 36);
            this.btnLimparVistaExplodida.TabIndex = 12;
            this.btnLimparVistaExplodida.BackColor = System.Drawing.Color.White;
            this.btnLimparVistaExplodida.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLimparVistaExplodida.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLimparVistaExplodida.Text = "Limpar explodida";
            this.btnLimparVistaExplodida.UseVisualStyleBackColor = true;
            this.btnLimparVistaExplodida.Click += new System.EventHandler(this.btnLimparVistaExplodida_Click);
            // 
            // lblStatusTekla
            // 
            this.lblStatusTekla.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(70)))), ((int)(((byte)(78)))), ((int)(((byte)(92)))));
            this.lblStatusTekla.Location = new System.Drawing.Point(22, 258);
            this.lblStatusTekla.Name = "lblStatusTekla";
            this.lblStatusTekla.Size = new System.Drawing.Size(276, 32);
            this.lblStatusTekla.TabIndex = 13;
            this.lblStatusTekla.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblStatusTekla.Visible = false;
            // 
            // btnToggleLog
            // 
            this.btnToggleLog.BackColor = System.Drawing.Color.White;
            this.btnToggleLog.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToggleLog.Location = new System.Drawing.Point(22, 294);
            this.btnToggleLog.Name = "btnToggleLog";
            this.btnToggleLog.Size = new System.Drawing.Size(276, 28);
            this.btnToggleLog.TabIndex = 14;
            this.btnToggleLog.Text = "+ Mostrar log";
            this.btnToggleLog.UseVisualStyleBackColor = true;
            this.btnToggleLog.Click += new System.EventHandler(this.btnToggleLog_Click);
            // 
            // pnlLog
            // 
            this.pnlLog.BackColor = System.Drawing.Color.White;
            this.pnlLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlLog.Controls.Add(this.txtSaida);
            this.pnlLog.Location = new System.Drawing.Point(22, 330);
            this.pnlLog.Name = "pnlLog";
            this.pnlLog.Size = new System.Drawing.Size(276, 174);
            this.pnlLog.TabIndex = 15;
            this.pnlLog.Visible = false;
            // 
            // txtSaida
            // 
            this.txtSaida.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtSaida.BackColor = System.Drawing.Color.White;
            this.txtSaida.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSaida.Location = new System.Drawing.Point(0, 0);
            this.txtSaida.Multiline = true;
            this.txtSaida.Name = "txtSaida";
            this.txtSaida.ReadOnly = true;
            this.txtSaida.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSaida.Size = new System.Drawing.Size(274, 172);
            this.txtSaida.TabIndex = 0;
            this.txtSaida.WordWrap = false;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(246)))), ((int)(((byte)(248)))));
            this.ClientSize = new System.Drawing.Size(320, 336);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Controls.Add(this.pnlLog);
            this.Controls.Add(this.btnToggleLog);
            this.Controls.Add(this.lblStatusTekla);
            this.Controls.Add(this.btnLimparVistaExplodida);
            this.Controls.Add(this.btnCriarAreaDefinida);
            this.Controls.Add(this.btnCriarFolhaInteira);
            this.Controls.Add(this.chkColorir);
            this.Controls.Add(this.chkLinhas);
            this.Controls.Add(this.chkGhostLinhas);
            this.Controls.Add(this.chkPlanoXY);
            this.Controls.Add(this.chkPlanoXZ);
            this.Controls.Add(this.chkPlanoZY);
            this.Controls.Add(this.lblDrawingStatusLink);
            this.Controls.Add(this.pnlDrawingIndicator);
            this.Controls.Add(this.lblTeklaStatusLink);
            this.Controls.Add(this.pnlTeklaIndicator);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Integracao Tekla";
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.pnlLog.ResumeLayout(false);
            this.pnlLog.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Panel pnlTeklaIndicator;
        private System.Windows.Forms.Label lblTeklaStatusLink;
        private System.Windows.Forms.Panel pnlDrawingIndicator;
        private System.Windows.Forms.Label lblDrawingStatusLink;
        private System.Windows.Forms.CheckBox chkPlanoZY;
        private System.Windows.Forms.CheckBox chkPlanoXZ;
        private System.Windows.Forms.CheckBox chkPlanoXY;
        private System.Windows.Forms.CheckBox chkGhostLinhas;
        private System.Windows.Forms.CheckBox chkLinhas;
        private System.Windows.Forms.CheckBox chkColorir;
        private System.Windows.Forms.Button btnCriarFolhaInteira;
        private System.Windows.Forms.Button btnCriarAreaDefinida;
        private System.Windows.Forms.Button btnLimparVistaExplodida;
        private System.Windows.Forms.Label lblStatusTekla;
        private System.Windows.Forms.Button btnToggleLog;
        private System.Windows.Forms.Panel pnlLog;
        private System.Windows.Forms.TextBox txtSaida;
    }
}
