using System;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace ElmahHelper.Tools
{
    public class FormControlTool
    {
        /// <summary>
        /// New Button
        /// </summary>
        /// <param name="btnName"></param>
        /// <param name="btnText"></param>
        /// <param name="top"></param>
        /// <param name="left"></param>
        /// <returns></returns>
        public Button NewButton(string btnName, string btnText, int top, int left)
        {
            Button button = new Button();
            button.AutoSize = true;
            button.Name = btnName;
            button.Text = btnText;
            button.Top = top;
            button.Left = left;

            return button;
        }

        /// <summary>
        /// 打開SelectFolder視窗
        /// </summary>
        /// <param name="defaultPath">預設資料夾</param>
        /// <returns>選擇資料夾的路徑</returns>
        public string GetSelectFolderPath(string defaultPath = "")
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "請選擇一個資料夾";
                fbd.SelectedPath = defaultPath;     // 預設開啟的資料夾

                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    return fbd.SelectedPath;
                }

                return defaultPath;
            }
        }

        /// <summary>
        /// New DataGridView
        /// </summary>
        /// <param name="dataGridViewName"></param>
        /// <returns></returns>
        public DataGridView NewDataGridView(string dataGridViewName)
        {
            DataGridView dataGridView = new DataGridView();
            dataGridView.Name = dataGridViewName;
            dataGridView.AutoSize = true;
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.AutoGenerateColumns = true;
            //根據欄位內容調整欄位長度
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            return dataGridView;
        }

        /// <summary>
        /// NewMenuStrip
        /// </summary>
        /// <param name="menuStripName"></param>
        /// <returns></returns>
        public MenuStrip NewMenuStrip(string menuStripName)
        {
            MenuStrip menuStrip = new MenuStrip();
            menuStrip.Name = menuStripName;
            menuStrip.AutoSize = true;
            menuStrip.Dock = DockStyle.Top;

            return menuStrip;
        }

        /// <summary>
        /// NewToolStripMenuItem
        /// </summary>
        /// <param name="toolStripMenuItemName"></param>
        /// <param name="toolStripMenuItemText"></param>
        /// <param name="clickHandler"></param>
        /// <returns></returns>
        public ToolStripMenuItem NewToolStripMenuItem(string toolStripMenuItemName, string toolStripMenuItemText, EventHandler clickHandler = null)
        {
            ToolStripMenuItem toolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem.Name = toolStripMenuItemName;
            toolStripMenuItem.Text = toolStripMenuItemText;
            if( clickHandler != null )
                toolStripMenuItem.Click += clickHandler;

            return toolStripMenuItem;
        }

        /// <summary>
        /// NewToolStripMenuItemDropDownList
        /// </summary>
        /// <param name="toolStripMenuItemName"></param>
        /// <param name="toolStripMenuItemText"></param>
        /// <param name="dorpDownList"></param>
        /// <returns></returns>
        public ToolStripMenuItem NewToolStripMenuItemDropDownList(string toolStripMenuItemName, string toolStripMenuItemText, ToolStripMenuItem[] dorpDownList)
        {
            ToolStripMenuItem dropdownList = NewToolStripMenuItem(toolStripMenuItemName, toolStripMenuItemText);

            dropdownList.DropDownItems.AddRange(dorpDownList);

            return dropdownList;
        }

        /// <summary>
        /// NewDataGridViewButtonColumn
        /// </summary>
        /// <param name="dataGridViewButtonColumnName"></param>
        /// <param name="headerText"></param>
        /// <param name="dataGridViewButtonColumnText"></param>
        /// <returns></returns>
        public DataGridViewButtonColumn NewDataGridViewButtonColumn(string dataGridViewButtonColumnName, string headerText, string dataGridViewButtonColumnText)
        {
            DataGridViewButtonColumn btnCol = new DataGridViewButtonColumn();
            btnCol.Name = dataGridViewButtonColumnName;
            btnCol.HeaderText = headerText;
            btnCol.Text = dataGridViewButtonColumnText;
            //true：
            //    每個按鈕格子都會顯示buttonColumn.Text的值
            //    簡單說：所有列的按鈕都會顯示一樣的文字
            //false：
            //    系統會從該儲存格的值(cell.Value)來顯示文字
            //    可以針對不同的列，設定不同的按鈕文字
            btnCol.UseColumnTextForButtonValue = true;

            return btnCol;
        }

        /// <summary>
        /// 生成一個DataGridView的ButtonColumn
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataGridView"></param>
        /// <param name="dataGridViewButtonColumnName"></param>
        /// <param name="actionColName"></param>
        /// <param name="btnText"></param>
        /// <param name="callback"></param>
        public void GenDataGridViewActionColumn<T>(DataGridView dataGridView, string dataGridViewButtonColumnName, string actionColName, string btnText, Action<T> callback)
        {
            DataGridViewButtonColumn btnCol = NewDataGridViewButtonColumn(dataGridViewButtonColumnName, actionColName, btnText);
            dataGridView.Columns.Add(btnCol);

            dataGridView.CellClick += (sender, e) =>
            {
                var dgv = sender as DataGridView;

                // 確保不是點到標題列或空白列
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

                // 檢查是否是你按鈕的那一欄
                if (dgv.Columns[e.ColumnIndex].Name == dataGridViewButtonColumnName)
                    //抓出這列綁定的資料物件並執行 callback
                    if (dgv.Rows[e.RowIndex].DataBoundItem is T item)
                        callback?.Invoke(item);
            };
        }
    }
}
