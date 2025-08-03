using ElmahHelper.Service;
using ElmahHelper.Tools;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ElmahHelper.Model
{
    public class ElmahPage
    {
        private ElmahService elmahSrv = new ElmahService();
        private FormControlTool controlTool = new FormControlTool();
        public IList<Elmah> ElmahList { get; set; }
        public ElmahQueryCondition ElmahQueryCondition { get; set; }
        public DataGridView ElmahDataGridView { get; set; }
        public Form DetailForm { get; set; }
        public TabPage ElmahTabPage { get; set; }

        /// <summary>
        /// 查詢Elmah
        /// </summary>
        /// <remarks>一併更新_elmahList</remarks>
        public void QueryError()
        {
            //LoadElmahList
            ElmahList = elmahSrv.GetElmahList(ElmahQueryCondition.ElmahSourceFolderPath
                                             , ElmahQueryCondition.StartTime
                                             , ElmahQueryCondition.EndTime
                                             , ElmahQueryCondition.FileName
                                             , ElmahQueryCondition.Message
                                             , ElmahQueryCondition.Detail);

            ElmahDataGridView.DataSource =  new BindingList<Error>(ElmahList.Select(elmah => elmah.ElmahError).ToList());
        }

        /// <summary>
        /// 刪除Grid上的Elmah
        /// </summary>
        /// <remarks>如檔案在zip裡，會刪除整個zip。會備份至BackUp\yyyyMMdd-HHmmss</remarks>
        public void DeleteElmah()
        {
            elmahSrv.DeleteElmah((IList<Error>)ElmahDataGridView.DataSource, ElmahList);

            QueryError();
        }

        /// <summary>
        /// 打開Detail視窗
        /// </summary>
        /// <param name="error"></param>
        public void OpenErrorDetail(Error error)
        {
            DetailForm.Controls["ErrorDetailTextBox"].Text = error.GetDetail();

            DetailForm.ShowDialog(); // 模態顯示
        }

        /// <summary>
        /// 在檔案總管中開啟
        /// </summary>
        /// <param name="error"></param>
        public void OpenElmahFolder(Error error)
        {
            Elmah selectedElmah = ElmahList.Where(elmah => elmah.GUID == error.ErrorID).FirstOrDefault();

            if (string.IsNullOrEmpty(selectedElmah.SourceZIPPath))
                Process.Start("explorer.exe", $"/select,\"{Path.Combine(selectedElmah.ParentFolderPath, selectedElmah.FileName)}\"");
            else
                Process.Start("explorer.exe", $"/select,\"{Path.Combine(selectedElmah.SourceZIPPath, selectedElmah.FileName)}\"");
        }

        /// <summary>
        /// 初始化
        /// </summary>
        public ElmahPage()
        {
            ElmahList = new List<Elmah>();
            
            DetailForm = GenErrorDetailForm();

            ElmahQueryCondition = new ElmahQueryCondition();

            ElmahDataGridView = GenErrorDataGridView();

            QueryError();

            ElmahTabPage = GenQueryElmahTabPage();
        }

        #region 生成畫面
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
        /// <returns></returns>
        private TableLayoutPanel GenQueryElmahTabPageLayout()
        {
            TableLayoutPanel queryElmahTabPageLayout = controlTool.NewTableLayoutPanel("QueryElmahTabPageLayout", 4, 1);

            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 0: 資訊區
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 1: 查詢區(Time)
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 2: 查詢區(Contain)
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     //Row 3: 動作區
            queryElmahTabPageLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); //Row 4: Grid區

            queryElmahTabPageLayout.Controls.Add(GenElmahInfoPanel(ElmahQueryCondition), 1, 0);    //資訊區
            queryElmahTabPageLayout.Controls.Add(GenElmahTimeQueryPanel(ElmahQueryCondition), 1, 1);    //查詢區(Time)
            queryElmahTabPageLayout.Controls.Add(GenElmahContainQueryPanel(ElmahQueryCondition), 1, 2);    //查詢區(Contain)
            queryElmahTabPageLayout.Controls.Add(GenElmahActionPanel(), 1, 3);    //動作區
            queryElmahTabPageLayout.Controls.Add(ElmahDataGridView, 1, 4);    //Grid區

            return queryElmahTabPageLayout;
        }

        /// <summary>
        /// 生成Elmah資訊區域
        /// </summary>
        /// <returns></returns>
        private Panel GenElmahInfoPanel(ElmahQueryCondition queryCondition)
        {
            int currLeft = 0;
            Panel infoPanel = new Panel
            {
                Name = "ElmahInfoPanel"
                , Dock = DockStyle.Fill
                , Height = 40 // 或其他適合高度
            };

            queryCondition.ElmahSourceFolderPathLabel.Location = new Point(currLeft, 15);

            infoPanel.Controls.Add(queryCondition.ElmahSourceFolderPathLabel);

            return infoPanel;
        }

        /// <summary>
        /// 生成查詢區域(Time)
        /// </summary>
        /// <remarks>查詢區間:[StartDateTime] ~ [EndDateTime]</remarks>
        /// <returns></returns>
        private Panel GenElmahTimeQueryPanel(ElmahQueryCondition queryCondition)
        {
            int currLeft = 0;
    
            Panel queryPanel = new Panel
            {
                Name = "ElmahTimeQueryPanel"
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

            queryCondition.StartTimePicker.Location = new Point(currLeft, 10);
            currLeft += queryCondition.StartTimePicker.Width;
            queryPanel.Controls.Add(queryCondition.StartTimePicker);

            lable = new Label 
            { 
                Text = "~"
                , Location = new Point(currLeft + 10, 15)
                , Width = 20 
            };
            currLeft += lable.Width + 10;
            queryPanel.Controls.Add(lable);

            queryCondition.EndTimePicker.Location = new Point(currLeft, 10);
            currLeft += queryCondition.EndTimePicker.Width + 10;
            queryPanel.Controls.Add(queryCondition.EndTimePicker);

            return queryPanel;
        }

        /// <summary>
        /// 生成查詢區域(Contain)
        /// </summary>
        /// <remarks>檔名:[FileNameTextBox] Message:[MessageTextBox] Detail:[DetailTextBox]</remarks>
        /// <returns></returns>
        private Panel GenElmahContainQueryPanel(ElmahQueryCondition queryCondition)
        {
            int currLeft = 0;
            Panel queryPanel = new Panel
            {
                Name = "ElmahContainQueryPanel"
                , Dock = DockStyle.Fill
                , Height = 40 // 或其他適合高度
            };

            GenElmahContainQueryItem(queryPanel, ref currLeft, "檔名:" , 35, queryCondition.FileNameTextBox);
            GenElmahContainQueryItem(queryPanel, ref currLeft, "Message:" , 50, queryCondition.MessageTextBox);
            GenElmahContainQueryItem(queryPanel, ref currLeft, "Detail:" , 45, queryCondition.DetailTextBox);

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
        private void GenElmahContainQueryItem(Panel queryPanel, ref int currLeft, string labelText, int labelWidth, TextBox textBox)
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
        private Panel GenElmahActionPanel()
        {
            int currLeft = 0;

            Panel actionPanel = new Panel
            {
                Name = "ElmahActionPanel"
                , Dock = DockStyle.Fill
                , Height = 40
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
            currLeft += deleteBtn.Width + 10;
            actionPanel.Controls.Add(deleteBtn);

            Button changeElmahFolderBtn = new Button
            {
                Text = "更改Elmah資料夾"
                , Location = new Point(currLeft + 10, 10)
                , Width = 150
            };
            changeElmahFolderBtn.Click += ElmahQueryCondition.ChangeElmahFolder;
            currLeft += changeElmahFolderBtn.Width + 10;
            actionPanel.Controls.Add(changeElmahFolderBtn);

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
            dataGridView.DataBindingComplete += (sender, e) => { GenGridAction(sender, e); };

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
                , (error) => { OpenErrorDetail(error); });
            }

            if (!dataGridView.Columns.Contains("OpenElmahFolderCol"))
            {
                controlTool.GenDataGridViewActionColumn<Error>(dataGridView
                , "OpenElmahFolderCol"
                , "操作", "檔案總管顯示"
                , 0
                , (error) => { OpenElmahFolder(error); });
            }
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
    }
}
