using ProcessExcels.Bal;
using ProcessExcels.Dal;
using ProcessExcels.Model;
using ProcessExcels.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProcessExcels
{
    public partial class frmProcessExcel : Form
    {
        #region Local Variables
        const string CLASS_NAME = "frmProcessExcel";
        ILogger logger;
        #endregion

        #region Constructor
        public frmProcessExcel()
        {
            InitializeComponent();
            logger = new Logger();

            lblStatus.ForeColor = Color.Green;
            btnProcess.Enabled = false;
        }
        #endregion

        #region Private Events
        private void btnBrowseExcelFiles_Click(object sender, EventArgs e)
        {
            try
            {
                btnProcess.Enabled = false;
                tabExcelControl.TabPages.Clear();

                string[] filePaths = browseFiles();

                if (filePaths != null)
                {
                    loadExcelFilesIntoDataGridViews(filePaths);
                }
            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "btnBrowseExcelFiles_Click", ex);
            }
        }
        private void btnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                string tabPageName = getSelctedTabName();

                List<Dictionary<string, object>> valuePairs = processGridViewRecords(tabPageName);

                if (valuePairs.Count > 0)
                {
                    lblStatus.Text = "Please wait while the files being processed....";

                    addRecordsToDatabase(tabPageName, valuePairs);

                    lblStatus.Text = "Records processed and inserted successfully into database!";

                    askAndSendEmail();

                    lblStatus.Text = "Records processed and inserted successfully into database and email sent.!";

                }
            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "btnProcess_Click", ex);
            }
        }
        private void tabExcelControl_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            List<string> finalFileList = new List<string>();

            foreach (string file in files)
            {

                if (file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    finalFileList.Add(file);
                }

                if (finalFileList.Count > 0)
                {
                    loadExcelFilesIntoDataGridViews(finalFileList.ToArray());
                }
            }
        }
        private void tabExcelControl_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region Private Methods
        private void askAndSendEmail()
        {
            DialogResult dialogResult = MessageBox.Show("Records processed and inserted successfully into database.\r\nDo you want to send an email alert for loaded records ?",
                                        "Send Email Alert",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question);

            if (dialogResult == DialogResult.Yes)
            {

                sendEmail();
            }

        }
        private void sendEmail()
        {
            try
            {
                ISendEmail sendEmail = new SendEmail()
                {
                    SmtpServer = "smtp.gmail.com",
                    SmtpPort = 587,
                    SmtpUserName = "<provide proper email>",
                    SmtpPassword = "<provide proper password>",
                    SenderEmail = "<provide proper email>",
                    RecipientEmail = "<provide proper email>",
                    Subject = "Email alert for successfully loaded records.",
                    Body = "",
                    AttachmentPath = Common.GetLogFilePath()
                };

                // Below api will work with "Less secure access enable for gmail"

                //sendEmail.Send();
                
                lblStatus.Text = "Records processed and inserted successfully into database and email sent.!";
            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "sendEmail", ex);
                throw;
            }
        }
        private string[] browseFiles()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Select Excel files";
                openFileDialog.InitialDirectory = @"C:\";
                openFileDialog.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
                openFileDialog.Multiselect = true;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileNames;
                }
                else
                {
                    MessageBox.Show("No excel files selected..", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }
            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "browseFiles", ex);

                return null;
            }

        }
        private string getSelctedTabName()
        {
            if (tabExcelControl.SelectedTab != null)
            {
                return tabExcelControl.SelectedTab.AccessibilityObject.Name;
            }
            return string.Empty;
        }
        private async void loadExcelFilesIntoDataGridViews(string[] filePaths)
        {
            try
            {
                lblStatus.Text = "Please wait while the files being processed....";
                List<RecordsLoadResult> resultRecords = new List<RecordsLoadResult>();

                foreach (string filepath in filePaths)
                {
                    var result = await getRecordsFromSheet(filepath);
                    resultRecords.Add(result);
                }

                foreach (RecordsLoadResult record in resultRecords)
                {
                    DataTable table = new DataTable();
                    table = record.DataTable;
                    string tabPageName = record.TabPageName;

                    DataGridView dataGridView = new DataGridView();
                    dataGridView.Dock = DockStyle.Fill;
                    dataGridView.DataSource = table;
                    dataGridView.ReadOnly = true;

                    TabPage tabPage = new TabPage(tabPageName);
                    tabPage.Controls.Add(dataGridView);
                    tabExcelControl.TabPages.Add(tabPage);

                    logger.LogInfo($" Total Records: '{table.Rows.Count}' from file :'{tabPageName}' loaded successfully.");

                    lblStatus.Text = "Records loaded successfully!";
                    btnProcess.Enabled = true;
                }

            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "loadExcelFilesIntoDataGridViews", ex);
                throw;
            }
        }
        private async Task<RecordsLoadResult> getRecordsFromSheet(string filepath)
        {
            try
            {
                return await Task.Run(() =>
                {
                    IExcelOperations excelOperations = new ExcelOperations();
                    string[] sheetsList = excelOperations.GetAllSheetsFromExcelFile(filepath);
                    string firstSheetName = sheetsList[0];

                    DataTable dataTable = excelOperations.GetRecordsFromSheet(filepath, firstSheetName);

                    Func<string, string> getFileNameFromPath = fname => Path.GetFileNameWithoutExtension(Path.GetFileName(fname));

                    string pagetab = getFileNameFromPath(filepath) + "_" + firstSheetName.Trim('$');

                    RecordsLoadResult recordsLoadResult = new RecordsLoadResult()
                    {
                        TabPageName = pagetab,
                        DataTable = dataTable
                    };

                    return recordsLoadResult;
                });
            }
            catch (Exception)
            {
                throw;
            }
        }
        private List<Dictionary<string, object>> processGridViewRecords(string tabPageName)
        {
            try
            {
                TabPage tabPage = tabExcelControl.SelectedTab;
                List<Dictionary<string, object>> valuePairs = new List<Dictionary<string, object>>();

                if (tabPage != null)
                {
                    foreach (Control control in tabPage.Controls)
                    {
                        if (control is DataGridView gridView)
                        {
                            foreach (DataGridViewRow row in gridView.SelectedRows)
                            {
                                var gridRowColValDic = new Dictionary<string, object>();

                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    string columnName = gridView.Columns[cell.ColumnIndex].Name.Trim().Replace(" ", "_");
                                    object columnValue = cell.Value;
                                    gridRowColValDic.Add(columnName, columnValue);
                                }
                                valuePairs.Add(gridRowColValDic);
                            }
                        }
                    }
                }
                return valuePairs;
            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "processGridViewRecords", ex);
                throw;
            }
        }
        private async void addRecordsToDatabase(string tableName, List<Dictionary<string, object>> valuePairs)
        {
            try
            {
                await Task.Run(() =>
                {
                    foreach (Dictionary<string, object> valuePair in valuePairs)
                    {
                        IProcessDataDal processData = new ProcessDataDal();

                        if (!processData.TableAvailableInDataBase(tableName))
                        {
                            string createTableQuery = processData.BuildCreateTableQuery(tableName, valuePair);
                            processData.CreateTabel(createTableQuery);
                        }

                        string inserquesry = processData.BuildInsertQuery(tableName, valuePair);
                        processData.InsertRecords(inserquesry);
                    }

                });
            }
            catch (Exception ex)
            {
                logger.LogError(CLASS_NAME, "addRecordsToDatabase", ex);
                throw;
            }
        }
        #endregion
    }
}
