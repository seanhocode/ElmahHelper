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

        private string _elmahsourceFolderPath;
        private ElmahQueryCondition _elmahQueryCondition;
        private IList<Elmah> _elmahList;
        private DataGridView _errorDataGridView;
        private Form _detailForm;

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
            // Step.1 _elmahsourceFolderPath
            _elmahsourceFolderPath = ConfigurationManager.AppSettings["DefaultElmahFolderPath"];

            // Step.2 _elmahQueryCondition
            GenElmahQueryCondition();

            // Step.3 _errorDataGridView
            _errorDataGridView = GenErrorDataGridView();

            // Step.4 _detailForm
            _detailForm = GenErrorDetailForm();

            // Step.5 _elmahList
            QueryError();
        }
        #endregion
        #region Elmah
        #region 生成畫面
        /// <summary>
        /// 初始化ElmahQueryCondition
        /// </summary>
        private void GenElmahQueryCondition()
        {
            _elmahQueryCondition = new ElmahQueryCondition();
            string dateTimeFormat = "yyyy/MM/dd HH:mm";

            _elmahQueryCondition.StartTimePicker =
                controlTool.NewDateTimePicker("QueryCondition_StartTime", DateTime.Today);
            _elmahQueryCondition.StartTimePicker.CustomFormat = dateTimeFormat;

            _elmahQueryCondition.EndTimePicker =
                controlTool.NewDateTimePicker("QueryCondition_EndTime", DateTime.Today.AddDays(1));
            _elmahQueryCondition.EndTimePicker.CustomFormat = dateTimeFormat;

            _elmahQueryCondition.FileNameTextBox = controlTool.NewTextBox("QueryCondition_FileName", 120);
            _elmahQueryCondition.MessageTextBox = controlTool.NewTextBox("QueryCondition_Message", 120);
            _elmahQueryCondition.DetailTextBox = controlTool.NewTextBox("QueryCondition_Detail", 120);
        }

        /// <summary>
        /// 生成主要畫面
        /// </summary>
        /// <returns></returns>
        private TableLayoutPanel GenMainLayout()
        {
            TableLayoutPanel mainLayout = controlTool.NewTableLayoutPanel("MainLayout", 2, 1);

            // =======================================MainLayout=======================================
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      //Row 0: 菜單
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  //Row 1: MainTabControl

            mainLayout.Controls.Add(GenMainMenuStrip()  , 1, 0);            //菜單
            mainLayout.Controls.Add(GenMainTabControl() , 1, 1);            //MainTabControl
            // =======================================MainLayout=======================================

            return mainLayout;
        }

        /// <summary>
        /// 生成主畫面TabControl
        /// </summary>
        /// <returns></returns>
        private TabControl GenMainTabControl()
        {
            TabControl mainTabControl = controlTool.NewTabControl("MainTabControl");

            mainTabControl.Controls.Add(GenQueryElmahTabPage());

            return mainTabControl;
        }

        /// <summary>
        /// 生成查詢Elmah頁籤
        /// </summary>
        /// <returns></returns>
        private TabPage GenQueryElmahTabPage()
        {
            TabPage queryElmahTabPage = controlTool.NewTabPage("QueryElmahTabPage", "查詢");

            queryElmahTabPage.Controls.Add(GenQueryElmahTabPageLayout());

            return queryElmahTabPage;
        }

        /// <summary>
        /// 生成查詢Elmah頁籤內容
        /// </summary>
        /// <remarks>需先Load _errorDataGridView</remarks>
        /// <returns></returns>
        private TableLayoutPanel GenQueryElmahTabPageLayout()
        {
            TableLayoutPanel queryElmahTabPageLayout = controlTool.NewTableLayoutPanel("QueryElmahTabPageLayout", 4, 1);

            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 0: 查詢區(Time)
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 1: 查詢區(Contain)
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 2: 動作區
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); //Row 3: Grid區

            queryElmahTabPageLayout.Controls.Add(GenTimeQueryPanel()    , 1, 0);        //查詢區(Time)
            queryElmahTabPageLayout.Controls.Add(GenContainQueryPanel() , 1, 1);        //查詢區(Contain)
            queryElmahTabPageLayout.Controls.Add(GenActionPanel()       , 1, 2);        //動作區
            queryElmahTabPageLayout.Controls.Add(_errorDataGridView     , 1, 3);        //Grid區

            return queryElmahTabPageLayout;
        }

        /// <summary>
        /// 生成查詢區域(Time)
        /// </summary>
        /// <remarks>查詢區間:[StartDateTime] ~ [EndDateTime]</remarks>
        /// <returns></returns>
        private Panel GenTimeQueryPanel()
        {
            int currLeft = 0;
            
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

            _elmahQueryCondition.StartTimePicker.Location = new Point(currLeft, 10);
            currLeft += _elmahQueryCondition.StartTimePicker.Width;
            queryPanel.Controls.Add(_elmahQueryCondition.StartTimePicker);

            lable = new Label 
            { 
                Text = "~"
                , Location = new Point(currLeft + 10, 15)
                , Width = 20 
            };
            currLeft += lable.Width + 10;
            queryPanel.Controls.Add(lable);

            _elmahQueryCondition.EndTimePicker.Location = new Point(currLeft, 10);
            currLeft += _elmahQueryCondition.EndTimePicker.Width + 10;
            queryPanel.Controls.Add(_elmahQueryCondition.EndTimePicker);

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

            GenContainQueryItem(queryPanel, ref currLeft, "檔名:" , 35, _elmahQueryCondition.FileNameTextBox);
            GenContainQueryItem(queryPanel, ref currLeft, "Message:" , 50, _elmahQueryCondition.MessageTextBox);
            GenContainQueryItem(queryPanel, ref currLeft, "Detail:" , 45, _elmahQueryCondition.DetailTextBox);

            return queryPanel;
        }

        /// <summary>
        /// 生成ContainQueryItem
        /// </summary>
        /// <param name="queryPanel"></param>
        /// <param name="currLeft"></param>
        /// <param name="labelText"></param>
        /// <param name="labelWidth"></param>
        /// <param name="textBox"></param>
        /// <remarks>[Label][TextBox]</remarks>
        private void GenContainQueryItem(Panel queryPanel, ref int currLeft, string labelText, int labelWidth, TextBox textBox)
        {
            Label lable = new Label
            {
                Text = labelText
                , Location = new Point(currLeft, 15)
                , Width = labelWidth
            };
            currLeft += lable.Width;
            queryPanel.Controls.Add(lable);

            textBox.Location = new Point(currLeft, 10);
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
        /// 生成ElmahGird區域
        /// </summary>
        /// <remarks>會執行QueryError()，因此會一併更新_elmahList</remarks>
        private DataGridView GenErrorDataGridView()
        {
            DataGridView dataGridView = controlTool.NewDataGridView("ErrorDataGridView");

            dataGridView.DataSource = new BindingList<Error>();

            //資料Binding完後生成Grid按鈕
            dataGridView.DataBindingComplete += GenGridAction;

            return dataGridView;
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
                , 0
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
                            , saveElmahFolderItem = controlTool.NewToolStripMenuItem("SaveElmahFolderStripMenuItem", "儲存Elmah資料夾", SaveElmahFolder)
                            , fileDropDownList = controlTool.NewToolStripMenuItemDropDownList("FileToolStripMenuItem", "檔案", new ToolStripMenuItem[] { openElmahFolderItem, saveElmahFolderItem });

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
        /// <remarks>一併更新_elmahList</remarks>
        private void QueryError()
        {
            //LoadElmahList
            _elmahList = elmahSrv.GetElmahList(_elmahsourceFolderPath
                                             , _elmahQueryCondition.StartTime
                                             , _elmahQueryCondition.EndTime
                                             , _elmahQueryCondition.FileName
                                             , _elmahQueryCondition.Message
                                             , _elmahQueryCondition.Detail);

            _errorDataGridView.DataSource = new BindingList<Error>(_elmahList.Select(elmah => elmah.ElmahError).ToList());
        }

        /// <summary>
        /// 刪除Grid上的Elmah
        /// </summary>
        /// <remarks>如檔案在zip裡，會刪除整個zip。會備份至BackUp\yyyyMMdd-HHmmss</remarks>
        private void DeleteElmah()
        {
            elmahSrv.DeleteElmah((IList<Error>)_errorDataGridView.DataSource, _elmahList);

            QueryError();
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
        #endregion
    }
}
