using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ElmahHelper.Tools
{
    public class FileTool
    {
        /// <summary>
        /// 目前應用程式的基底目錄，執行檔所在的資料夾
        /// </summary>
        public static readonly string CurrentFolder = AppContext.BaseDirectory;

        /// <summary>
        /// 檢查檔案是否存在
        /// </summary>
        /// <param name="filePath">檔案完整路徑</param>
        /// <returns>檔案是否存在</returns>
        /// <exception cref="ArgumentNullException">filePath為null或空字串時拋出</exception>
        public bool CheckFileExist(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath), "檔案路徑不可為空");

            return File.Exists(filePath);
        }

        /// <summary>
        /// 檢查資料夾是否存在
        /// </summary>
        /// <param name="folderPath">資料夾完整路徑</param>
        /// <param name="isAutoCreate">資料夾不存在時是否自動建立</param>
        /// <returns>結束後資料夾是否存在</returns>
        /// <exception cref="ArgumentNullException">folderPath為null或空白字串時拋出</exception>
        public bool CheckFolderExist(string folderPath, bool isAutoCreate = false)
        {

            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentNullException(nameof(folderPath), "資料夾路徑不可為空");

            if (Directory.Exists(folderPath))
                return true;

            if (isAutoCreate)
            {
                Directory.CreateDirectory(folderPath);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 取得指定資料夾內所有檔案的完整路徑清單。
        /// </summary>
        /// <param name="folderPath">資料夾完整路徑</param>
        /// <returns>該資料夾底下的所有檔案(不包含子資料夾)</returns>
        public List<string> GetAllFileInFolder(string folderPath)
        {
            if (!CheckFolderExist(folderPath))
                return null;

            return Directory.GetFiles(folderPath).ToList();
        }
    }
}
