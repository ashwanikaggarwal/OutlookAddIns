namespace OutlookAddIns.Forms.Controls
{
    partial class tableList
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lstFormFields = new System.Windows.Forms.ListBox();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.lblFields = new System.Windows.Forms.Label();
            this.lblTableName = new System.Windows.Forms.Label();
            this.lstDbFields = new System.Windows.Forms.ListBox();
            this.lblDbFields = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lstFormFields
            // 
            this.lstFormFields.FormattingEnabled = true;
            this.lstFormFields.Location = new System.Drawing.Point(15, 62);
            this.lstFormFields.Name = "lstFormFields";
            this.lstFormFields.Size = new System.Drawing.Size(155, 173);
            this.lstFormFields.TabIndex = 0;
            // 
            // txtTableName
            // 
            this.txtTableName.Location = new System.Drawing.Point(15, 23);
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(155, 20);
            this.txtTableName.TabIndex = 1;
            // 
            // lblFields
            // 
            this.lblFields.AutoSize = true;
            this.lblFields.Location = new System.Drawing.Point(12, 46);
            this.lblFields.Name = "lblFields";
            this.lblFields.Size = new System.Drawing.Size(60, 13);
            this.lblFields.TabIndex = 2;
            this.lblFields.Text = "Form Fields";
            // 
            // lblTableName
            // 
            this.lblTableName.AutoSize = true;
            this.lblTableName.Location = new System.Drawing.Point(12, 7);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(65, 13);
            this.lblTableName.TabIndex = 3;
            this.lblTableName.Text = "Table Name";
            // 
            // lstDbFields
            // 
            this.lstDbFields.FormattingEnabled = true;
            this.lstDbFields.Location = new System.Drawing.Point(186, 62);
            this.lstDbFields.Name = "lstDbFields";
            this.lstDbFields.Size = new System.Drawing.Size(155, 173);
            this.lstDbFields.TabIndex = 4;
            // 
            // lblDbFields
            // 
            this.lblDbFields.AutoSize = true;
            this.lblDbFields.Location = new System.Drawing.Point(183, 46);
            this.lblDbFields.Name = "lblDbFields";
            this.lblDbFields.Size = new System.Drawing.Size(78, 13);
            this.lblDbFields.TabIndex = 5;
            this.lblDbFields.Text = "Database Field";
            // 
            // tableList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.lblDbFields);
            this.Controls.Add(this.lstDbFields);
            this.Controls.Add(this.lblTableName);
            this.Controls.Add(this.lblFields);
            this.Controls.Add(this.txtTableName);
            this.Controls.Add(this.lstFormFields);
            this.Name = "tableList";
            this.Size = new System.Drawing.Size(357, 252);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lstFormFields;
        private System.Windows.Forms.TextBox txtTableName;
        private System.Windows.Forms.Label lblFields;
        private System.Windows.Forms.Label lblTableName;
        private System.Windows.Forms.ListBox lstDbFields;
        private System.Windows.Forms.Label lblDbFields;
    }
}
