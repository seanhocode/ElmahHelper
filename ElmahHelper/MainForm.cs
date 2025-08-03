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
        private FormControlTool controlTool = new FormControlTool();

        //private ElmahPage _elmahPage;
        private IList<ElmahPage> _elmahPageList;
        private TabControl _mainTabControl;


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
            _elmahPageList = new List<ElmahPage>();
            _elmahPageList.Add(new ElmahPage());

            GenMainTabControl();
        }
        #endregion

        #region 生成畫面
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
            mainLayout.Controls.Add(_mainTabControl, 1, 1);                 //MainTabControl
            // =======================================MainLayout=======================================

            return mainLayout;
        }

        /// <summary>
        /// 生成主畫面TabControl
        /// </summary>
        /// <returns></returns>
        private void GenMainTabControl()
        {
            _mainTabControl = controlTool.NewTabControl("MainTabControl");

            _mainTabControl.Controls.Add(_elmahPageList[0].ElmahTabPage);
        }

        /// <summary>
        /// 生成主選單區域
        /// </summary>
        /// <returns></returns>
        private MenuStrip GenMainMenuStrip()
        {
            MenuStrip mainMenuStrip = controlTool.NewMenuStrip("MainMenuStrip");
            ToolStripMenuItem openElmahFolderItem = controlTool.NewToolStripMenuItem("OpenElmahFolderStripMenuItem", "新增查詢頁面", NewElmahQueryPage)
                            , fileDropDownList = controlTool.NewToolStripMenuItemDropDownList("FileToolStripMenuItem", "功能", new ToolStripMenuItem[] { openElmahFolderItem });

            mainMenuStrip.Items.Add(fileDropDownList);

            return mainMenuStrip;
        }
        #endregion

        /// <summary>
        /// 新增查詢Elmah頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewElmahQueryPage(object sender, EventArgs e)
        {
            ElmahPage page = new ElmahPage();
            page.ElmahQueryCondition.ChangeElmahFolder();

            _elmahPageList.Add(page);
            _mainTabControl.Controls.Add(page.ElmahTabPage);
        }
    }
}
