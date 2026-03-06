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
            this.chkGhostLinhas = new System.Windows.Forms.CheckBox();
            this.chkLinhas = new System.Windows.Forms.CheckBox();
            this.chkColorir = new System.Windows.Forms.CheckBox();
            this.chkXPositivo = new System.Windows.Forms.CheckBox();
            this.chkXNegativo = new System.Windows.Forms.CheckBox();
            this.chkYPositivo = new System.Windows.Forms.CheckBox();
            this.chkYNegativo = new System.Windows.Forms.CheckBox();
            this.chkZPositivo = new System.Windows.Forms.CheckBox();
            this.chkZNegativo = new System.Windows.Forms.CheckBox();
            this.btnCriarFolhaInteira = new System.Windows.Forms.Button();
            this.btnCriarAreaDefinida = new System.Windows.Forms.Button();
            this.btnLimparVistaExplodida = new System.Windows.Forms.Button();
            this.lblStatusTekla = new System.Windows.Forms.Label();
            this.btnToggleLog = new System.Windows.Forms.Button();
            this.chkTeste = new System.Windows.Forms.CheckBox();
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
            // chkGhostLinhas
            // 
            this.chkGhostLinhas.AutoSize = true;
            this.chkGhostLinhas.Checked = true;
            this.chkGhostLinhas.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkGhostLinhas.Location = new System.Drawing.Point(22, 58);
            this.chkGhostLinhas.Name = "chkGhostLinhas";
            this.chkGhostLinhas.Size = new System.Drawing.Size(64, 20);
            this.chkGhostLinhas.TabIndex = 4;
            this.chkGhostLinhas.Text = "Ghost";
            this.chkGhostLinhas.UseVisualStyleBackColor = true;
            // 
            // chkLinhas
            // 
            this.chkLinhas.AutoSize = true;
            this.chkLinhas.Checked = true;
            this.chkLinhas.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLinhas.Location = new System.Drawing.Point(115, 58);
            this.chkLinhas.Name = "chkLinhas";
            this.chkLinhas.Size = new System.Drawing.Size(67, 20);
            this.chkLinhas.TabIndex = 5;
            this.chkLinhas.Text = "Linhas";
            this.chkLinhas.UseVisualStyleBackColor = true;
            // 
            // chkColorir
            // 
            this.chkColorir.AutoSize = true;
            this.chkColorir.Location = new System.Drawing.Point(208, 58);
            this.chkColorir.Name = "chkColorir";
            this.chkColorir.Size = new System.Drawing.Size(68, 20);
            this.chkColorir.TabIndex = 6;
            this.chkColorir.Text = "Colorir";
            this.chkColorir.UseVisualStyleBackColor = true;
            // 
            // chkXPositivo
            // 
            this.chkXPositivo.AutoSize = true;
            this.chkXPositivo.Checked = true;
            this.chkXPositivo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkXPositivo.ForeColor = System.Drawing.Color.Firebrick;
            this.chkXPositivo.Location = new System.Drawing.Point(22, 88);
            this.chkXPositivo.Name = "chkXPositivo";
            this.chkXPositivo.Size = new System.Drawing.Size(42, 20);
            this.chkXPositivo.TabIndex = 7;
            this.chkXPositivo.Text = "X+";
            this.chkXPositivo.UseVisualStyleBackColor = true;
            // 
            // chkXNegativo
            // 
            this.chkXNegativo.AutoSize = true;
            this.chkXNegativo.Checked = true;
            this.chkXNegativo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkXNegativo.ForeColor = System.Drawing.Color.Firebrick;
            this.chkXNegativo.Location = new System.Drawing.Point(91, 88);
            this.chkXNegativo.Name = "chkXNegativo";
            this.chkXNegativo.Size = new System.Drawing.Size(40, 20);
            this.chkXNegativo.TabIndex = 8;
            this.chkXNegativo.Text = "X-";
            this.chkXNegativo.UseVisualStyleBackColor = true;
            // 
            // chkYPositivo
            // 
            this.chkYPositivo.AutoSize = true;
            this.chkYPositivo.Checked = true;
            this.chkYPositivo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkYPositivo.ForeColor = System.Drawing.Color.ForestGreen;
            this.chkYPositivo.Location = new System.Drawing.Point(158, 88);
            this.chkYPositivo.Name = "chkYPositivo";
            this.chkYPositivo.Size = new System.Drawing.Size(41, 20);
            this.chkYPositivo.TabIndex = 9;
            this.chkYPositivo.Text = "Y+";
            this.chkYPositivo.UseVisualStyleBackColor = true;
            // 
            // chkYNegativo
            // 
            this.chkYNegativo.AutoSize = true;
            this.chkYNegativo.Checked = true;
            this.chkYNegativo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkYNegativo.ForeColor = System.Drawing.Color.ForestGreen;
            this.chkYNegativo.Location = new System.Drawing.Point(225, 88);
            this.chkYNegativo.Name = "chkYNegativo";
            this.chkYNegativo.Size = new System.Drawing.Size(39, 20);
            this.chkYNegativo.TabIndex = 10;
            this.chkYNegativo.Text = "Y-";
            this.chkYNegativo.UseVisualStyleBackColor = true;
            // 
            // chkZPositivo
            // 
            this.chkZPositivo.AutoSize = true;
            this.chkZPositivo.Checked = true;
            this.chkZPositivo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkZPositivo.ForeColor = System.Drawing.Color.RoyalBlue;
            this.chkZPositivo.Location = new System.Drawing.Point(22, 114);
            this.chkZPositivo.Name = "chkZPositivo";
            this.chkZPositivo.Size = new System.Drawing.Size(42, 20);
            this.chkZPositivo.TabIndex = 11;
            this.chkZPositivo.Text = "Z+";
            this.chkZPositivo.UseVisualStyleBackColor = true;
            // 
            // chkZNegativo
            // 
            this.chkZNegativo.AutoSize = true;
            this.chkZNegativo.Checked = true;
            this.chkZNegativo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkZNegativo.ForeColor = System.Drawing.Color.RoyalBlue;
            this.chkZNegativo.Location = new System.Drawing.Point(91, 114);
            this.chkZNegativo.Name = "chkZNegativo";
            this.chkZNegativo.Size = new System.Drawing.Size(40, 20);
            this.chkZNegativo.TabIndex = 12;
            this.chkZNegativo.Text = "Z-";
            this.chkZNegativo.UseVisualStyleBackColor = true;
            // 
            // btnCriarFolhaInteira
            // 
            this.btnCriarFolhaInteira.Location = new System.Drawing.Point(22, 150);
            this.btnCriarFolhaInteira.Name = "btnCriarFolhaInteira";
            this.btnCriarFolhaInteira.Size = new System.Drawing.Size(276, 36);
            this.btnCriarFolhaInteira.TabIndex = 13;
            this.btnCriarFolhaInteira.BackColor = System.Drawing.Color.White;
            this.btnCriarFolhaInteira.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCriarFolhaInteira.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCriarFolhaInteira.Text = "Explodir em folha inteira";
            this.btnCriarFolhaInteira.UseVisualStyleBackColor = true;
            this.btnCriarFolhaInteira.Click += new System.EventHandler(this.btnCriarFolhaInteira_Click);
            // 
            // btnCriarAreaDefinida
            // 
            this.btnCriarAreaDefinida.Location = new System.Drawing.Point(22, 194);
            this.btnCriarAreaDefinida.Name = "btnCriarAreaDefinida";
            this.btnCriarAreaDefinida.Size = new System.Drawing.Size(276, 36);
            this.btnCriarAreaDefinida.TabIndex = 14;
            this.btnCriarAreaDefinida.BackColor = System.Drawing.Color.White;
            this.btnCriarAreaDefinida.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCriarAreaDefinida.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCriarAreaDefinida.Text = "Definir area";
            this.btnCriarAreaDefinida.UseVisualStyleBackColor = true;
            this.btnCriarAreaDefinida.Click += new System.EventHandler(this.btnCriarAreaDefinida_Click);
            // 
            // btnLimparVistaExplodida
            // 
            this.btnLimparVistaExplodida.Location = new System.Drawing.Point(22, 238);
            this.btnLimparVistaExplodida.Name = "btnLimparVistaExplodida";
            this.btnLimparVistaExplodida.Size = new System.Drawing.Size(276, 36);
            this.btnLimparVistaExplodida.TabIndex = 15;
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
            this.lblStatusTekla.Location = new System.Drawing.Point(22, 282);
            this.lblStatusTekla.Name = "lblStatusTekla";
            this.lblStatusTekla.Size = new System.Drawing.Size(276, 32);
            this.lblStatusTekla.TabIndex = 16;
            this.lblStatusTekla.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblStatusTekla.Visible = false;
            // 
            // btnToggleLog
            // 
            this.btnToggleLog.BackColor = System.Drawing.Color.White;
            this.btnToggleLog.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToggleLog.Location = new System.Drawing.Point(22, 318);
            this.btnToggleLog.Name = "btnToggleLog";
            this.btnToggleLog.Size = new System.Drawing.Size(276, 28);
            this.btnToggleLog.TabIndex = 17;
            this.btnToggleLog.Text = "+ Mostrar log";
            this.btnToggleLog.UseVisualStyleBackColor = true;
            this.btnToggleLog.Click += new System.EventHandler(this.btnToggleLog_Click);
            // 
            // chkTeste
            // 
            this.chkTeste.AutoSize = true;
            this.chkTeste.Location = new System.Drawing.Point(22, 353);
            this.chkTeste.Name = "chkTeste";
            this.chkTeste.Size = new System.Drawing.Size(60, 20);
            this.chkTeste.TabIndex = 18;
            this.chkTeste.Text = "Teste";
            this.chkTeste.UseVisualStyleBackColor = true;
            // 
            // pnlLog
            // 
            this.pnlLog.BackColor = System.Drawing.Color.White;
            this.pnlLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlLog.Controls.Add(this.txtSaida);
            this.pnlLog.Location = new System.Drawing.Point(22, 381);
            this.pnlLog.Name = "pnlLog";
            this.pnlLog.Size = new System.Drawing.Size(276, 174);
            this.pnlLog.TabIndex = 19;
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
            this.ClientSize = new System.Drawing.Size(320, 387);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Controls.Add(this.pnlLog);
            this.Controls.Add(this.chkTeste);
            this.Controls.Add(this.btnToggleLog);
            this.Controls.Add(this.lblStatusTekla);
            this.Controls.Add(this.btnLimparVistaExplodida);
            this.Controls.Add(this.btnCriarAreaDefinida);
            this.Controls.Add(this.btnCriarFolhaInteira);
            this.Controls.Add(this.chkZNegativo);
            this.Controls.Add(this.chkZPositivo);
            this.Controls.Add(this.chkYNegativo);
            this.Controls.Add(this.chkYPositivo);
            this.Controls.Add(this.chkXNegativo);
            this.Controls.Add(this.chkXPositivo);
            this.Controls.Add(this.chkColorir);
            this.Controls.Add(this.chkLinhas);
            this.Controls.Add(this.chkGhostLinhas);
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
        private System.Windows.Forms.CheckBox chkGhostLinhas;
        private System.Windows.Forms.CheckBox chkLinhas;
        private System.Windows.Forms.CheckBox chkColorir;
        private System.Windows.Forms.CheckBox chkXPositivo;
        private System.Windows.Forms.CheckBox chkXNegativo;
        private System.Windows.Forms.CheckBox chkYPositivo;
        private System.Windows.Forms.CheckBox chkYNegativo;
        private System.Windows.Forms.CheckBox chkZPositivo;
        private System.Windows.Forms.CheckBox chkZNegativo;
        private System.Windows.Forms.Button btnCriarFolhaInteira;
        private System.Windows.Forms.Button btnCriarAreaDefinida;
        private System.Windows.Forms.Button btnLimparVistaExplodida;
        private System.Windows.Forms.Label lblStatusTekla;
        private System.Windows.Forms.Button btnToggleLog;
        private System.Windows.Forms.CheckBox chkTeste;
        private System.Windows.Forms.Panel pnlLog;
        private System.Windows.Forms.TextBox txtSaida;
    }
}
