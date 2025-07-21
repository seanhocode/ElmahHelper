using ElmahHelper.Model;
using ElmahHelper.Service;
using ElmahHelper.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ElmahHelper
{
    public partial class MainForm: Form
    {
        private ElmahService elmahSrv = new ElmahService();
        private FormControlTool controlTool = new FormControlTool();
        private FileTool fileTool = new FileTool();

        private IList<Elmah> _elmahList;
        private DataGridView _errorDataGridView;
        private string _elmahsourceFolderPath;
        private Form _detailForm;

        private DateTime _errorStartTime;
        private DateTime _errorEndTime;
        private string _fileNameContain;
        private string _messageContain;
        private string _detailContain;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            InitialData();

            this.Controls.Add(GenMainLayout());
        }

        #region Data
        private void InitialData()
        {
            _elmahsourceFolderPath = ConfigurationManager.AppSettings["DefaultElmahFolderPath"];
            _errorStartTime = DateTime.Today;//Today:日期, Now:日期+時間
            _errorEndTime = DateTime.Today.AddDays(1);
            _fileNameContain = string.Empty;
            _messageContain = string.Empty;
            _detailContain = string.Empty;
            _errorDataGridView = controlTool.NewDataGridView("ErrorDataGridView");
            _elmahList = elmahSrv.GetElmahList(_elmahsourceFolderPath, _errorStartTime, _errorEndTime);
            _detailForm = GenErrorDetailForm();
        }

        /// <summary>
        /// 更新Error Grid的資料
        /// </summary>
        private void LoadErrorDataGridView()
        {
            _elmahList = elmahSrv.GetElmahList(_elmahsourceFolderPath, _errorStartTime, _errorEndTime, _fileNameContain, _messageContain, _detailContain);
            _errorDataGridView.DataSource = new BindingList<Error>(_elmahList.Select(elmah => elmah.ElmahError).ToList());
        }
        #endregion

        #region 生成畫面
        /// <summary>
        /// 生成主要畫面
        /// </summary>
        /// <returns></returns>
        private TableLayoutPanel GenMainLayout()
        {
            GenErrorDataGridView(_errorDataGridView, _elmahList.Select(elmah => elmah.ElmahError).ToList());

            TableLayoutPanel mainLayout = controlTool.NewTableLayoutPanel("MainLayout", 2, 1)
                           , queryElmahTabPageLayout = controlTool.NewTableLayoutPanel("QueryElmahTabPageLayout", 4, 1);

            TabControl mainTabControl = controlTool.NewTabControl("MainTabControl");
            TabPage queryElmahTabPage = controlTool.NewTabPage("QueryElmahTabPage", "查詢");

            // =======================================QueryElmahTabPageLayout==========================
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      //Row 0: 查詢區(Time)
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      //Row 1: 查詢區(Contain)
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      //Row 2: 動作區
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  //Row 3: Grid區

            queryElmahTabPageLayout.Controls.Add(GenTimeQueryPanel()    , 1, 0);    //查詢區(Time)
            queryElmahTabPageLayout.Controls.Add(GenContainQueryPanel() , 1, 1);    //查詢區(Contain)
            queryElmahTabPageLayout.Controls.Add(GenActionPanel(), 1, 2);    //動作區
            queryElmahTabPageLayout.Controls.Add(_errorDataGridView     , 1, 3);    //Grid區

            queryElmahTabPage.Controls.Add(queryElmahTabPageLayout);
            // =======================================QueryElmahTabPageLayout==========================

            // =======================================MainTabControl===================================
            mainTabControl.Controls.Add(queryElmahTabPage);
            // =======================================MainTabControl===================================

            // =======================================MainLayout=======================================
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      //Row 0: 菜單
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  //Row 1: MainTabControl

            mainLayout.Controls.Add(GenMainMenuStrip()  , 1, 0);    //菜單
            mainLayout.Controls.Add(mainTabControl      , 1, 1);    //MainTabControl
            // =======================================MainLayout=======================================

            return mainLayout;
        }

        /// <summary>
        /// 生成查詢區域(Time)
        /// </summary>
        /// <remarks>查詢區間:[StartDateTime] ~ [EndDateTime]</remarks>
        /// <returns></returns>
        private Panel GenTimeQueryPanel()
        {
            int currLeft = 0;
            string dateTimeFormat = "yyyy/MM/dd HH:mm";
            Panel queryPanel = new Panel
            {
                Name = "TimeQueryPanel"
                , Dock = DockStyle.Fill
                , Height = 40 // 或其他適合高度
            };

            Label lable = new Label
            {
                Text = "查詢區間:"
                , Location = new Point(currLeft, 15)
                , Width = 60
            };
            currLeft += lable.Width;
            queryPanel.Controls.Add(lable);

            // DateTimePicker 起始
            DateTimePicker startTime = controlTool.NewDateTimePicker("QueryCondition_StartTime", currLeft, 10, _errorStartTime);
            startTime.CustomFormat = dateTimeFormat;
            currLeft += startTime.Width;
            queryPanel.Controls.Add(startTime);

            lable = new Label 
            { 
                Text = "~"
                , Location = new Point(currLeft + 10, 15)
                , Width = 20 
            };
            currLeft += lable.Width + 10;
            queryPanel.Controls.Add(lable);

            // DateTimePicker 結束
            DateTimePicker endTime = controlTool.NewDateTimePicker("QueryCondition_EndTime", currLeft, 10, _errorEndTime);
            endTime.CustomFormat = dateTimeFormat;
            currLeft += endTime.Width + 10;
            queryPanel.Controls.Add(endTime);

            return queryPanel;
        }

        /// <summary>
        /// 生成查詢區域(Contain)
        /// </summary>
        /// <remarks>檔名:[FileNameTextBox] Message:[MessageTextBox] Detail:[DetailTextBox]</remarks>
        /// <returns></returns>
        private Panel GenContainQueryPanel()
        {
            int currLeft = 0;
            Panel queryPanel = new Panel
            {
                Name = "ContainQueryPanel"
                , Dock = DockStyle.Fill
                , Height = 40 // 或其他適合高度
            };

            GenContainQueryItem(queryPanel, ref currLeft, "檔名:" , 35, "QueryCondition_FileName");
            GenContainQueryItem(queryPanel, ref currLeft, "Message:" , 50, "QueryCondition_Message");
            GenContainQueryItem(queryPanel, ref currLeft, "Detail:" , 45, "QueryCondition_Detail");

            return queryPanel;
        }

        /// <summary>
        /// 生成ContainQueryItem
        /// </summary>
        /// <param name="queryPanel"></param>
        /// <param name="currLeft"></param>
        /// <param name="labelText"></param>
        /// <param name="labelWidth"></param>
        /// <param name="textBoxName"></param>
        /// <remarks>[Label][TextBox]</remarks>
        private void GenContainQueryItem(Panel queryPanel, ref int currLeft, string labelText, int labelWidth, string textBoxName)
        {
            Label lable = new Label
            {
                Text = labelText
                , Location = new Point(currLeft, 15)
                , Width = labelWidth
            };
            currLeft += lable.Width;
            queryPanel.Controls.Add(lable);

            TextBox textBox = new TextBox
            {
                Name = textBoxName
                , Location = new Point(currLeft, 10)
                , Width = 120
            };
            currLeft += textBox.Width;
            queryPanel.Controls.Add(textBox);
        }

        /// <summary>
        /// 生成動作區域區域
        /// </summary>
        /// <remarks>[查詢Btn] [刪除Btn]</remarks>
        /// <returns></returns>
        private Panel GenActionPanel()
        {
            int currLeft = 0;

            Panel actionPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 40 // 或其他適合高度
            };

            Button queryBtn = new Button
            {
                Text = "查詢"
                , Location = new Point(currLeft + 10, 10)
                , Width = 50
            };
            queryBtn.Click += (sender, e) =>
            { QueryError(); MessageBox.Show("Done!"); };
            currLeft += queryBtn.Width + 10;
            actionPanel.Controls.Add(queryBtn);

            Button deleteBtn = new Button
            {
                Text = "刪除"
                , Location = new Point(currLeft + 10, 10)
                , Width = 50
            };
            deleteBtn.Click += (sender, e) =>
            { DeleteElmah(); };
            actionPanel.Controls.Add(deleteBtn);

            return actionPanel;
        }

        /// <summary>
        /// 生成Gird區域
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="errorList"></param>
        private void GenErrorDataGridView(DataGridView dataGridView, IList<Error> errorList)
        {
            dataGridView.DataSource = new BindingList<Error>(errorList);

            dataGridView.DataBindingComplete += GenGridAction;
        }

        /// <summary>
        /// 生成Grid的動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenGridAction(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            DataGridView dataGridView = (DataGridView)sender;

            if (!dataGridView.Columns.Contains("OpenErrorDetailCol"))
            {
                controlTool.GenDataGridViewActionColumn<Error>(dataGridView
                , "OpenErrorDetailCol"
                , "操作", "細節"
                , 0
                , (error) =>
                {
                    OpenErrorDetail(error);
                });
            }

            if (!dataGridView.Columns.Contains("OpenElmahFolderCol"))
            {
                controlTool.GenDataGridViewActionColumn<Error>(dataGridView
                , "OpenElmahFolderCol"
                , "操作", "檔案總管顯示"
                , 1
                , (error) =>
                {
                    OpenElmahFolder(error);
                });
            }
        }

        /// <summary>
        /// 生成主選單區域
        /// </summary>
        /// <returns></returns>
        private MenuStrip GenMainMenuStrip()
        {
            MenuStrip mainMenuStrip = controlTool.NewMenuStrip("MainMenuStrip");
            ToolStripMenuItem openElmahFolderItem = controlTool.NewToolStripMenuItem("OpenElmahFolderStripMenuItem", "打開Elmah資料夾", LoadElmahFolder)
                            , saveElmahFolderItem = controlTool.NewToolStripMenuItem("SaveElmahFolderStripMenuItem", "儲存Elmah資料夾", SaveElmahFolder);
            ToolStripMenuItem fileDropDownList = controlTool.NewToolStripMenuItemDropDownList("FileToolStripMenuItem", "檔案", new ToolStripMenuItem[] { openElmahFolderItem, saveElmahFolderItem });

            mainMenuStrip.Items.Add(fileDropDownList);

            return mainMenuStrip;
        }

        /// <summary>
        /// 生成Detail Form
        /// </summary>
        /// <returns></returns>
        public Form GenErrorDetailForm()
        {
            Form errorDetailForm = new Form
            {
                Dock = DockStyle.Fill
                , Text = "Detail"
                , AutoScroll = true
                , Size = new Size(1200, 600)
            };

            errorDetailForm.FormClosing += (s, e) =>
            {
                e.Cancel = true;    // 不關閉
                ((Form)s).Hide();   // 改為隱藏
            };

            GenErrorDetailFormArea(errorDetailForm);

            return errorDetailForm;
        }

        /// <summary>
        /// 生成DetailForm內容
        /// </summary>
        /// <param name="errorDetailForm"></param>
        private void GenErrorDetailFormArea(Form errorDetailForm)
        {
            TextBox textBox = new TextBox
            {
                Name = "ErrorDetailTextBox"
                , Multiline = true
                , Dock = DockStyle.Fill
                , ScrollBars = ScrollBars.Both
                , AutoSize = true
            };

            errorDetailForm.Controls.Add(textBox);
        }
        #endregion

        #region ButtonClick
        /// <summary>
        /// 載入ElmahFolder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoadElmahFolder(object sender, EventArgs e)
        {
            _elmahsourceFolderPath = controlTool.GetSelectFolderPath(_elmahsourceFolderPath);

            QueryError();
        }

        /// <summary>
        /// 儲存ElmahFolder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveElmahFolder(object sender, EventArgs e)
        {
            Configuration config =
                ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings["DefaultElmahFolderPath"].Value = _elmahsourceFolderPath;

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            MessageBox.Show("Done!");
        }

        /// <summary>
        /// 查詢Elmah
        /// </summary>
        private void QueryError()
        {
            TableLayoutPanel tableLayout = (TableLayoutPanel)Controls["MainLayout"];
            TabControl tabControl = (TabControl)tableLayout.Controls["MainTabControl"];
            TabPage tabPage = (TabPage)tabControl.Controls["QueryElmahTabPage"];
            tableLayout = (TableLayoutPanel)tabPage.Controls["QueryElmahTabPageLayout"];
            Panel panel = (Panel)tableLayout.Controls["TimeQueryPanel"];

            
            DateTimePicker dateTimePicker = (DateTimePicker)panel.Controls["QueryCondition_StartTime"];
            _errorStartTime = dateTimePicker.Value;
            dateTimePicker = (DateTimePicker)panel.Controls["QueryCondition_EndTime"];
            _errorEndTime = dateTimePicker.Value;

            panel = (Panel)tableLayout.Controls["ContainQueryPanel"];
            TextBox textBox = (TextBox)panel.Controls["QueryCondition_FileName"];
            _fileNameContain = textBox.Text;
            textBox = (TextBox)panel.Controls["QueryCondition_Message"];
            _messageContain = textBox.Text;
            textBox = (TextBox)panel.Controls["QueryCondition_Detail"];
            _detailContain = textBox.Text;

            LoadErrorDataGridView();
        }

        /// <summary>
        /// 刪除Grid上的Elmah
        /// </summary>
        /// <remarks>如檔案在zip裡，會刪除整個zip。會備份至BackUp\yyyyMMdd-HHmmss</remarks>
        private void DeleteElmah()
        {
            string filePathMsg = string.Empty
                 , backupFolderPath = Path.Combine(FileTool.CurrentFolder, "BackUp", $"{DateTime.Now.ToString("yyyyMMdd-HHmmss")}");

            //Step.1 取得Grid上的Elmah
            IList<Error> gridErrorList = (IList<Error>)_errorDataGridView.DataSource;
            IList<Elmah> deleteElmahList = _elmahList.Where(elmah => gridErrorList.Any(error => error.ErrorID.Contains(elmah.GUID))).ToList();

            //Step.2 取得類型不為zip的Elmah(直接刪除)
            List<string> deleteFilePath = deleteElmahList
                                                .Where(elmah => string.IsNullOrEmpty(elmah.SourceZIPPath))
                                                .Select(elmah => Path.Combine(elmah.ParentFolderPath, elmah.FileName))
                                                .Distinct()
                                                .ToList();

            //Step.3 取得類型為zip的Elmah(刪除zip)
            deleteFilePath.AddRange(deleteElmahList
                                                .Where(elmah => !string.IsNullOrEmpty(elmah.SourceZIPPath))
                                                .Select(elmah => elmah.SourceZIPPath)
                                                .Distinct()
                                                .ToList());
            //Step.4 取得要刪除的檔案，通知使用者
            foreach (string filePath in deleteFilePath)
                filePathMsg += $"{Path.GetFileName(filePath)}\r\n";

            if (filePathMsg != string.Empty && controlTool.OpenYesNoForm("確認", $"是否刪除以下檔案?\r\n(如檔案在zip裡，會刪除整個zip)\r\n{filePathMsg}"))
            {
                fileTool.CheckFolderExist(backupFolderPath, true);
                foreach (string filePath in deleteFilePath)
                {
                    //Step.5 備份、刪除檔案(直接移至輩分資料夾)
                    File.Move(filePath, Path.Combine(backupFolderPath, Path.GetFileName(filePath)));
                }

                MessageBox.Show("Done!");

                QueryError();
            }
            else
            {
                MessageBox.Show("無要刪除的檔案");
            }
        }

        /// <summary>
        /// 打開Detail視窗
        /// </summary>
        /// <param name="error"></param>
        private void OpenErrorDetail(Error error)
        {
            _detailForm.Controls["ErrorDetailTextBox"].Text = error.GetDetail();

            _detailForm.ShowDialog(); // 模態顯示
        }

        /// <summary>
        /// 在檔案總管中開啟
        /// </summary>
        /// <param name="error"></param>
        private void OpenElmahFolder(Error error)
        {
            Elmah selectedElmah = _elmahList.Where(elmah => elmah.GUID == error.ErrorID).FirstOrDefault();

            if(string.IsNullOrEmpty(selectedElmah.SourceZIPPath))
                Process.Start("explorer.exe", $"/select,\"{Path.Combine(selectedElmah.ParentFolderPath, selectedElmah.FileName)}\"");
            else
                Process.Start("explorer.exe", $"/select,\"{Path.Combine(selectedElmah.SourceZIPPath, selectedElmah.FileName)}\"");
        }
        #endregion
    }
}
