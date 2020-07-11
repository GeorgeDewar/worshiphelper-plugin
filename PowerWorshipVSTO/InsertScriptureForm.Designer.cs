namespace PowerWorshipVSTO
{
    partial class InsertScriptureForm
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
            this.txtBook = new System.Windows.Forms.TextBox();
            this.lblSearchBox = new System.Windows.Forms.Label();
            this.btnInsert = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtReference = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbTranslation = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // txtBook
            // 
            this.txtBook.Location = new System.Drawing.Point(89, 42);
            this.txtBook.Name = "txtBook";
            this.txtBook.Size = new System.Drawing.Size(198, 20);
            this.txtBook.TabIndex = 0;
            this.txtBook.TextChanged += new System.EventHandler(this.txtSearchBox_TextChanged);
            this.txtBook.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSearchBox_KeyPress);
            // 
            // lblSearchBox
            // 
            this.lblSearchBox.AutoSize = true;
            this.lblSearchBox.Location = new System.Drawing.Point(48, 45);
            this.lblSearchBox.Name = "lblSearchBox";
            this.lblSearchBox.Size = new System.Drawing.Size(35, 13);
            this.lblSearchBox.TabIndex = 1;
            this.lblSearchBox.Text = "Book:";
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(89, 94);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(95, 22);
            this.btnInsert.TabIndex = 2;
            this.btnInsert.Text = "Insert";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 71);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Reference:";
            // 
            // txtReference
            // 
            this.txtReference.Location = new System.Drawing.Point(89, 68);
            this.txtReference.Name = "txtReference";
            this.txtReference.Size = new System.Drawing.Size(198, 20);
            this.txtReference.TabIndex = 3;
            this.txtReference.TextChanged += new System.EventHandler(this.txtReference_TextChanged);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(190, 93);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(97, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Translation:";
            // 
            // cmbTranslation
            // 
            this.cmbTranslation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTranslation.FormattingEnabled = true;
            this.cmbTranslation.Location = new System.Drawing.Point(89, 15);
            this.cmbTranslation.Name = "cmbTranslation";
            this.cmbTranslation.Size = new System.Drawing.Size(198, 21);
            this.cmbTranslation.TabIndex = 8;
            this.cmbTranslation.SelectionChangeCommitted += new System.EventHandler(this.cmbTranslation_SelectionChangeCommitted);
            // 
            // InsertScriptureForm
            // 
            this.AcceptButton = this.btnInsert;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(350, 131);
            this.Controls.Add(this.cmbTranslation);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtReference);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.lblSearchBox);
            this.Controls.Add(this.txtBook);
            this.Name = "InsertScriptureForm";
            this.Text = "Insert Scripture";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtBook;
        private System.Windows.Forms.Label lblSearchBox;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtReference;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbTranslation;
    }
}