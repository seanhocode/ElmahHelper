using ElmahHelper.Model;
using ElmahHelper.Service;
using ElmahHelper.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ElmahHelper
{
    public partial class MainForm: Form
    {
        private ElmahService elmahSrv = new ElmahService();
        private FormControlTool controlTool = new FormControlTool();

        private IList<Elmah> _elmahList;
        private DataGridView _errorDataGridView;
        private DateTime _errorStartTime;
        private DateTime _errorEndTime;
        private string _elmahsourceFolderPath;
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
            _elmahsourceFolderPath = ConfigurationManager.AppSettings["DefaultElmahFolderPath"];
            _errorStartTime = DateTime.Today;//Today:日期, Now:日期+時間
            _errorEndTime = DateTime.Today.AddDays(1);
            _errorDataGridView = controlTool.NewDataGridView("ErrorDataGridView");
            _elmahList = elmahSrv.GetElmahList(_elmahsourceFolderPath, _errorStartTime, _errorEndTime);
            _detailForm = GenErrorDetailForm();
        }

        /// <summary>
        /// 更新Error Grid的資料
        /// </summary>
        private void LoadErrorDataGridView()
        {
            _elmahList = elmahSrv.GetElmahList(_elmahsourceFolderPath, _errorStartTime, _errorEndTime);
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

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill
                ,
                RowCount = 3
                ,
                ColumnCount = 1
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Row 0: 查詢區
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // Row 1: MainMenuStrip
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Row 2: DataGridView

            layout.Controls.Add(GenMainMenuStrip(), 1, 0); //菜單
            layout.Controls.Add(GenQueryPanel(), 1, 1);//查詢區域
            layout.Controls.Add(_errorDataGridView, 1, 2);//Grid區域

            return layout;
        }

        /// <summary>
        /// 生成查詢區域
        /// </summary>
        /// <remarks>查詢區間:[StartDateTime] ~ [EndDateTime] [查詢Btn]</remarks>
        /// <returns></returns>
        private Panel GenQueryPanel()
        {
            int currLeft = 0;
            string dateTimeFormat = "yyyy/MM/dd HH:mm";
            Panel queryPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 40 // 或其他適合高度
            };

            // DateTimePicker 起始
            DateTimePicker dateFrom = new DateTimePicker
            {
                Name = "ErrorTimeStartDateTimePicker"
                , Format = DateTimePickerFormat.Custom
                , Location = new Point(currLeft, 10)
                //, ShowUpDown = true
                , Value = _errorStartTime
            };
            dateFrom.CustomFormat = dateTimeFormat;
            currLeft += dateFrom.Width;
            queryPanel.Controls.Add(dateFrom);

            Label lable = new Label 
            { 
                Text = "~"
                , Location = new Point(currLeft + 10, 10)
                , Width = 20 
            };
            currLeft += lable.Width + 10;
            queryPanel.Controls.Add(lable);

            // DateTimePicker 結束
            DateTimePicker dateTo = new DateTimePicker
            {
                Name = "ErrorTimeEndDateTimePicker"
                , Format = DateTimePickerFormat.Custom
                , Location = new Point(currLeft, 10)
                //, ShowUpDown = true
                , Value = _errorEndTime
            };
            dateTo.CustomFormat = dateTimeFormat;
            currLeft += dateTo.Width + 10;
            queryPanel.Controls.Add(dateTo);

            Button queryBtn = new Button
            { 
                Text = "查詢"
                , Location = new Point(currLeft + 10, 10)
                , Width = 50
            };
            queryBtn.Click += (sender, e) => 
            { QueryBtn_Click(dateFrom.Value, dateTo.Value); };
            queryPanel.Controls.Add(queryBtn);

            return queryPanel;
        }

        /// <summary>
        /// 生成Gird區域
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="errorList"></param>
        private void GenErrorDataGridView(DataGridView dataGridView, IList<Error> errorList)
        {
            dataGridView.DataSource = new BindingList<Error>(errorList);

            controlTool.GenDataGridViewActionColumn<Error>(dataGridView
                , "OpenErrorDetailCol"
                , "操作", "顯示細節"
                , (error) =>
                    {
                        OpenDetailBtn_Click(error);
                    });
        }

        /// <summary>
        /// 生成主選單區域
        /// </summary>
        /// <returns></returns>
        private MenuStrip GenMainMenuStrip()
        {
            MenuStrip mainMenuStrip = controlTool.NewMenuStrip("MainMenuStrip");
            ToolStripMenuItem openElmahFolderItem = controlTool.NewToolStripMenuItem("OpenElmahFolderStripMenuItem", "打開Elmah資料夾", OpenElmahFolderStripMenuItem_Click)
                            , saveElmahFolderItem = controlTool.NewToolStripMenuItem("SaveElmahFolderStripMenuItem", "儲存Elmah資料夾", SaveElmahFolderStripMenuItem_Click);
            ToolStripMenuItem fileDropDownList = controlTool.NewToolStripMenuItemDropDownList("FileToolStripMenuItem", "檔案", new ToolStripMenuItem[] { openElmahFolderItem, saveElmahFolderItem });

            mainMenuStrip.Items.Add(fileDropDownList);

            return mainMenuStrip;
        }

        public Form GenErrorDetailForm()
        {
            Form errorDetailForm = new Form
            {
                Dock = DockStyle.Fill
                , Text = "Detail"
                , AutoScroll = true
                , Size = new Size(800, 800)
            };

            errorDetailForm.FormClosing += (s, e) =>
            {
                e.Cancel = true;    // 不關閉
                ((Form)s).Hide();   // 改為隱藏
            };

            GenErrorDetailFormArea(errorDetailForm);

            return errorDetailForm;
        }

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
        private void OpenElmahFolderStripMenuItem_Click(object sender, EventArgs e)
        {
            _elmahsourceFolderPath = controlTool.GetSelectFolderPath(_elmahsourceFolderPath);

            LoadErrorDataGridView();
        }

        private void SaveElmahFolderStripMenuItem_Click(object sender, EventArgs e)
        {
            Configuration config =
                ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            config.AppSettings.Settings["DefaultElmahFolderPath"].Value = _elmahsourceFolderPath;

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            MessageBox.Show("Done!");
        }

        private void QueryBtn_Click(DateTime errorStartTime, DateTime errorEndTime)
        {
            _errorStartTime = errorStartTime;
            _errorEndTime = errorEndTime;

            LoadErrorDataGridView();
        }

        private void OpenDetailBtn_Click(Error error)
        {
            _detailForm.Controls["ErrorDetailTextBox"].Text = error.GetDetail();

            _detailForm.ShowDialog(); // 模態顯示
        }
        #endregion
    }
}
