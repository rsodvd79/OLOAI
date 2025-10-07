namespace OutlookAiAddIn.UI
{
    internal partial class AiTaskPane
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                CancelPendingRequest();
                components?.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblMode = new System.Windows.Forms.Label();
            this.cmbMode = new System.Windows.Forms.ComboBox();
            this.btnLoadContext = new System.Windows.Forms.Button();
            this.lblContext = new System.Windows.Forms.Label();
            this.txtEmailBody = new System.Windows.Forms.TextBox();
            this.lblNotes = new System.Windows.Forms.Label();
            this.txtNotes = new System.Windows.Forms.TextBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblSuggestions = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblMode
            // 
            this.lblMode.AutoSize = true;
            this.lblMode.Location = new System.Drawing.Point(12, 14);
            this.lblMode.Name = "lblMode";
            this.lblMode.Size = new System.Drawing.Size(92, 13);
            this.lblMode.TabIndex = 0;
            this.lblMode.Text = "Tipo di supporto";
            // 
            // cmbMode
            // 
            this.cmbMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMode.FormattingEnabled = true;
            this.cmbMode.Location = new System.Drawing.Point(15, 30);
            this.cmbMode.Name = "cmbMode";
            this.cmbMode.Size = new System.Drawing.Size(250, 21);
            this.cmbMode.TabIndex = 1;
            this.cmbMode.SelectedIndexChanged += new System.EventHandler(this.cmbMode_SelectedIndexChanged);
            // 
            // btnLoadContext
            // 
            this.btnLoadContext.Location = new System.Drawing.Point(275, 28);
            this.btnLoadContext.Name = "btnLoadContext";
            this.btnLoadContext.Size = new System.Drawing.Size(120, 25);
            this.btnLoadContext.TabIndex = 2;
            this.btnLoadContext.Text = "Carica email";
            this.btnLoadContext.UseVisualStyleBackColor = true;
            this.btnLoadContext.Click += new System.EventHandler(this.btnLoadContext_Click);
            // 
            // lblContext
            // 
            this.lblContext.AutoSize = true;
            this.lblContext.Location = new System.Drawing.Point(12, 64);
            this.lblContext.Name = "lblContext";
            this.lblContext.Size = new System.Drawing.Size(116, 13);
            this.lblContext.TabIndex = 3;
            this.lblContext.Text = "Testo email / selezione";
            // 
            // txtEmailBody
            // 
            this.txtEmailBody.AcceptsReturn = true;
            this.txtEmailBody.AcceptsTab = true;
            this.txtEmailBody.Location = new System.Drawing.Point(15, 80);
            this.txtEmailBody.Multiline = true;
            this.txtEmailBody.Name = "txtEmailBody";
            this.txtEmailBody.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtEmailBody.Size = new System.Drawing.Size(380, 120);
            this.txtEmailBody.TabIndex = 3;
            // 
            // lblNotes
            // 
            this.lblNotes.AutoSize = true;
            this.lblNotes.Location = new System.Drawing.Point(12, 210);
            this.lblNotes.Name = "lblNotes";
            this.lblNotes.Size = new System.Drawing.Size(122, 13);
            this.lblNotes.TabIndex = 5;
            this.lblNotes.Text = "Istruzioni aggiuntive (opz)";
            // 
            // txtNotes
            // 
            this.txtNotes.AcceptsReturn = true;
            this.txtNotes.AcceptsTab = true;
            this.txtNotes.Location = new System.Drawing.Point(15, 226);
            this.txtNotes.Multiline = true;
            this.txtNotes.Name = "txtNotes";
            this.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtNotes.Size = new System.Drawing.Size(380, 60);
            this.txtNotes.TabIndex = 4;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(15, 296);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(120, 28);
            this.btnGenerate.TabIndex = 5;
            this.btnGenerate.Text = "Genera suggerimento";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(151, 296);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(120, 28);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Annulla";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblSuggestions
            // 
            this.lblSuggestions.AutoSize = true;
            this.lblSuggestions.Location = new System.Drawing.Point(12, 332);
            this.lblSuggestions.Name = "lblSuggestions";
            this.lblSuggestions.Size = new System.Drawing.Size(72, 13);
            this.lblSuggestions.TabIndex = 9;
            this.lblSuggestions.Text = "Suggerimento";
            // 
            // txtOutput
            // 
            this.txtOutput.AcceptsReturn = true;
            this.txtOutput.AcceptsTab = true;
            this.txtOutput.Location = new System.Drawing.Point(15, 348);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ReadOnly = true;
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtOutput.Size = new System.Drawing.Size(380, 150);
            this.txtOutput.TabIndex = 7;
            this.txtOutput.TextChanged += new System.EventHandler(this.txtOutput_TextChanged);
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(15, 506);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(120, 28);
            this.btnInsert.TabIndex = 8;
            this.btnInsert.Text = "Inserisci in Outlook";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(151, 506);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(120, 28);
            this.btnCopy.TabIndex = 9;
            this.btnCopy.Text = "Copia testo";
            this.btnCopy.UseVisualStyleBackColor = true;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.Location = new System.Drawing.Point(15, 542);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(380, 30);
            this.lblStatus.TabIndex = 12;
            this.lblStatus.Text = "Pronto";
            // 
            // AiTaskPane
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.lblSuggestions);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.txtNotes);
            this.Controls.Add(this.lblNotes);
            this.Controls.Add(this.txtEmailBody);
            this.Controls.Add(this.lblContext);
            this.Controls.Add(this.btnLoadContext);
            this.Controls.Add(this.cmbMode);
            this.Controls.Add(this.lblMode);
            this.Name = "AiTaskPane";
            this.Size = new System.Drawing.Size(410, 580);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.Label lblMode;
        private System.Windows.Forms.ComboBox cmbMode;
        private System.Windows.Forms.Button btnLoadContext;
        private System.Windows.Forms.Label lblContext;
        private System.Windows.Forms.TextBox txtEmailBody;
        private System.Windows.Forms.Label lblNotes;
        private System.Windows.Forms.TextBox txtNotes;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblSuggestions;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.Label lblStatus;
    }
}
