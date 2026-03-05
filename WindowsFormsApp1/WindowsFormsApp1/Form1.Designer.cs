namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnVerificarTekla = new System.Windows.Forms.Button();
            this.btnVerificarDesenho = new System.Windows.Forms.Button();
            this.btnVerVistaSelecionada = new System.Windows.Forms.Button();
            this.btnGerarVistaExplodida = new System.Windows.Forms.Button();
            this.btnLimparVistaExplodida = new System.Windows.Forms.Button();
            this.btnExplodirPorPlanos = new System.Windows.Forms.Button();
            this.chkDetectaAutomatico = new System.Windows.Forms.CheckBox();
            this.chkColorir = new System.Windows.Forms.CheckBox();
            this.chkGhostLinhas = new System.Windows.Forms.CheckBox();
            this.lblStatusTekla = new System.Windows.Forms.Label();
            this.txtSaida = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnVerificarTekla
            // 
            this.btnVerificarTekla.Location = new System.Drawing.Point(26, 26);
            this.btnVerificarTekla.Name = "btnVerificarTekla";
            this.btnVerificarTekla.Size = new System.Drawing.Size(214, 40);
            this.btnVerificarTekla.TabIndex = 0;
            this.btnVerificarTekla.Text = "Verificar comunicacao Tekla";
            this.btnVerificarTekla.UseVisualStyleBackColor = true;
            this.btnVerificarTekla.Click += new System.EventHandler(this.btnVerificarTekla_Click);
            // 
            // btnVerificarDesenho
            // 
            this.btnVerificarDesenho.Location = new System.Drawing.Point(260, 26);
            this.btnVerificarDesenho.Name = "btnVerificarDesenho";
            this.btnVerificarDesenho.Size = new System.Drawing.Size(214, 40);
            this.btnVerificarDesenho.TabIndex = 1;
            this.btnVerificarDesenho.Text = "Ver desenho aberto";
            this.btnVerificarDesenho.UseVisualStyleBackColor = true;
            this.btnVerificarDesenho.Click += new System.EventHandler(this.btnVerificarDesenho_Click);
            // 
            // btnVerVistaSelecionada
            // 
            this.btnVerVistaSelecionada.Location = new System.Drawing.Point(494, 26);
            this.btnVerVistaSelecionada.Name = "btnVerVistaSelecionada";
            this.btnVerVistaSelecionada.Size = new System.Drawing.Size(233, 40);
            this.btnVerVistaSelecionada.TabIndex = 2;
            this.btnVerVistaSelecionada.Text = "Ver vista selecionada";
            this.btnVerVistaSelecionada.UseVisualStyleBackColor = true;
            this.btnVerVistaSelecionada.Click += new System.EventHandler(this.btnVerVistaSelecionada_Click);
            // 
            // btnGerarVistaExplodida
            // 
            this.btnGerarVistaExplodida.Location = new System.Drawing.Point(747, 26);
            this.btnGerarVistaExplodida.Name = "btnGerarVistaExplodida";
            this.btnGerarVistaExplodida.Size = new System.Drawing.Size(183, 40);
            this.btnGerarVistaExplodida.TabIndex = 3;
            this.btnGerarVistaExplodida.Text = "Recriar conjunto + elementos";
            this.btnGerarVistaExplodida.UseVisualStyleBackColor = true;
            this.btnGerarVistaExplodida.Click += new System.EventHandler(this.btnGerarVistaExplodida_Click);
            // 
            // btnLimparVistaExplodida
            // 
            this.btnLimparVistaExplodida.Location = new System.Drawing.Point(747, 72);
            this.btnLimparVistaExplodida.Name = "btnLimparVistaExplodida";
            this.btnLimparVistaExplodida.Size = new System.Drawing.Size(183, 32);
            this.btnLimparVistaExplodida.TabIndex = 4;
            this.btnLimparVistaExplodida.Text = "Limpar explodida";
            this.btnLimparVistaExplodida.UseVisualStyleBackColor = true;
            this.btnLimparVistaExplodida.Click += new System.EventHandler(this.btnLimparVistaExplodida_Click);
            // 
            // btnExplodirPorPlanos
            // 
            this.btnExplodirPorPlanos.Location = new System.Drawing.Point(541, 72);
            this.btnExplodirPorPlanos.Name = "btnExplodirPorPlanos";
            this.btnExplodirPorPlanos.Size = new System.Drawing.Size(186, 32);
            this.btnExplodirPorPlanos.TabIndex = 5;
            this.btnExplodirPorPlanos.Text = "Explodir por planos (novo)";
            this.btnExplodirPorPlanos.UseVisualStyleBackColor = true;
            this.btnExplodirPorPlanos.Click += new System.EventHandler(this.btnExplodirPorPlanos_Click);
            // 
            // chkDetectaAutomatico
            // 
            this.chkDetectaAutomatico.AutoSize = true;
            this.chkDetectaAutomatico.Checked = true;
            this.chkDetectaAutomatico.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDetectaAutomatico.Location = new System.Drawing.Point(26, 79);
            this.chkDetectaAutomatico.Name = "chkDetectaAutomatico";
            this.chkDetectaAutomatico.Size = new System.Drawing.Size(225, 20);
            this.chkDetectaAutomatico.TabIndex = 6;
            this.chkDetectaAutomatico.Text = "Detecta automaticamente o conjunto";
            this.chkDetectaAutomatico.UseVisualStyleBackColor = true;
            // 
            // chkColorir
            // 
            this.chkColorir.AutoSize = true;
            this.chkColorir.Location = new System.Drawing.Point(272, 79);
            this.chkColorir.Name = "chkColorir";
            this.chkColorir.Size = new System.Drawing.Size(67, 20);
            this.chkColorir.TabIndex = 7;
            this.chkColorir.Text = "Colorir";
            this.chkColorir.UseVisualStyleBackColor = true;
            // 
            // chkGhostLinhas
            // 
            this.chkGhostLinhas.AutoSize = true;
            this.chkGhostLinhas.Checked = true;
            this.chkGhostLinhas.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkGhostLinhas.Location = new System.Drawing.Point(354, 79);
            this.chkGhostLinhas.Name = "chkGhostLinhas";
            this.chkGhostLinhas.Size = new System.Drawing.Size(111, 20);
            this.chkGhostLinhas.TabIndex = 8;
            this.chkGhostLinhas.Text = "Ghost + linhas";
            this.chkGhostLinhas.UseVisualStyleBackColor = true;
            // 
            // lblStatusTekla
            // 
            this.lblStatusTekla.AutoSize = true;
            this.lblStatusTekla.Location = new System.Drawing.Point(23, 114);
            this.lblStatusTekla.Name = "lblStatusTekla";
            this.lblStatusTekla.Size = new System.Drawing.Size(85, 16);
            this.lblStatusTekla.TabIndex = 9;
            this.lblStatusTekla.Text = "Status: N/D";
            // 
            // txtSaida
            // 
            this.txtSaida.Location = new System.Drawing.Point(26, 142);
            this.txtSaida.Multiline = true;
            this.txtSaida.Name = "txtSaida";
            this.txtSaida.ReadOnly = true;
            this.txtSaida.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSaida.Size = new System.Drawing.Size(904, 388);
            this.txtSaida.TabIndex = 10;
            this.txtSaida.WordWrap = false;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(960, 560);
            this.Controls.Add(this.txtSaida);
            this.Controls.Add(this.lblStatusTekla);
            this.Controls.Add(this.chkGhostLinhas);
            this.Controls.Add(this.chkColorir);
            this.Controls.Add(this.chkDetectaAutomatico);
            this.Controls.Add(this.btnExplodirPorPlanos);
            this.Controls.Add(this.btnLimparVistaExplodida);
            this.Controls.Add(this.btnGerarVistaExplodida);
            this.Controls.Add(this.btnVerVistaSelecionada);
            this.Controls.Add(this.btnVerificarDesenho);
            this.Controls.Add(this.btnVerificarTekla);
            this.Text = "Integracao Tekla";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button btnVerificarTekla;
        private System.Windows.Forms.Button btnVerificarDesenho;
        private System.Windows.Forms.Button btnVerVistaSelecionada;
        private System.Windows.Forms.Button btnGerarVistaExplodida;
        private System.Windows.Forms.Button btnLimparVistaExplodida;
        private System.Windows.Forms.Button btnExplodirPorPlanos;
        private System.Windows.Forms.CheckBox chkDetectaAutomatico;
        private System.Windows.Forms.CheckBox chkColorir;
        private System.Windows.Forms.CheckBox chkGhostLinhas;
        private System.Windows.Forms.Label lblStatusTekla;
        private System.Windows.Forms.TextBox txtSaida;
    }
}

