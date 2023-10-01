
namespace ProcessExcels
{
    partial class frmProcessExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmProcessExcel));
            this.btnBrowseExcelFiles = new System.Windows.Forms.Button();
            this.gbContainer = new System.Windows.Forms.GroupBox();
            this.tabExcelControl = new System.Windows.Forms.TabControl();
            this.btnProcess = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.txtNotes = new System.Windows.Forms.TextBox();
            this.gbContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBrowseExcelFiles
            // 
            this.btnBrowseExcelFiles.Location = new System.Drawing.Point(38, 30);
            this.btnBrowseExcelFiles.Name = "btnBrowseExcelFiles";
            this.btnBrowseExcelFiles.Size = new System.Drawing.Size(114, 23);
            this.btnBrowseExcelFiles.TabIndex = 2;
            this.btnBrowseExcelFiles.Text = "Browse Excel Files";
            this.btnBrowseExcelFiles.UseVisualStyleBackColor = true;
            this.btnBrowseExcelFiles.Click += new System.EventHandler(this.btnBrowseExcelFiles_Click);
            // 
            // gbContainer
            // 
            this.gbContainer.Controls.Add(this.tabExcelControl);
            this.gbContainer.Location = new System.Drawing.Point(36, 58);
            this.gbContainer.Name = "gbContainer";
            this.gbContainer.Size = new System.Drawing.Size(739, 296);
            this.gbContainer.TabIndex = 3;
            this.gbContainer.TabStop = false;
            // 
            // tabExcelControl
            // 
            this.tabExcelControl.AllowDrop = true;
            this.tabExcelControl.Location = new System.Drawing.Point(6, 19);
            this.tabExcelControl.Name = "tabExcelControl";
            this.tabExcelControl.SelectedIndex = 0;
            this.tabExcelControl.Size = new System.Drawing.Size(727, 272);
            this.tabExcelControl.TabIndex = 4;
            this.tabExcelControl.DragDrop += new System.Windows.Forms.DragEventHandler(this.tabExcelControl_DragDrop);
            this.tabExcelControl.DragEnter += new System.Windows.Forms.DragEventHandler(this.tabExcelControl_DragEnter);
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(158, 30);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(114, 23);
            this.btnProcess.TabIndex = 4;
            this.btnProcess.Text = "Process Records";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.Location = new System.Drawing.Point(42, 368);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 15);
            this.lblStatus.TabIndex = 5;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(661, 360);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(114, 23);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // txtNotes
            // 
            this.txtNotes.BackColor = System.Drawing.SystemColors.Menu;
            this.txtNotes.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtNotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNotes.Location = new System.Drawing.Point(39, 392);
            this.txtNotes.Multiline = true;
            this.txtNotes.Name = "txtNotes";
            this.txtNotes.Size = new System.Drawing.Size(736, 49);
            this.txtNotes.TabIndex = 7;
            this.txtNotes.Text = "Note:\r\nPress \"Browse Excel File\" button or Drag and Drop excel file on window to " +
    "load excelfile.\r\nPress \"Process Records\" button to insert selected records to da" +
    "tabase.\r\n";
            // 
            // frmProcessExcel
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txtNotes);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.gbContainer);
            this.Controls.Add(this.btnBrowseExcelFiles);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmProcessExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Process Excel Sheets";
            this.gbContainer.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnBrowseExcelFiles;
        private System.Windows.Forms.GroupBox gbContainer;
        private System.Windows.Forms.TabControl tabExcelControl;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.TextBox txtNotes;
    }
}

