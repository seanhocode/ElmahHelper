using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;

namespace ElmahHelper.Tools
{
    public class ZipTool
    {
        private FileTool fileTool = new FileTool();

        /// <summary>
        /// 取得指定ZIP壓縮檔中所有檔案(與資料夾)清單
        /// </summary>
        /// <param name="zipPath">ZIP壓縮檔的完整路徑</param>
        /// <remarks>不進行解壓縮，只列出內部的檔案名稱(含相對路徑)</remarks>
        /// <returns>ZIP中所有檔案(與資料夾)的路徑清單，如"folder/file.txt"</returns>
        public IList<string> GetFileNameInZip(string zipPath)
        {
            IList<string> fileNameList = new List<string>();

            if (!CheckZipFile(zipPath))
                return fileNameList;

            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    fileNameList.Add(entry.FullName);
                }
            }

            return fileNameList;
        }

        /// <summary>
        /// 檢查指定的檔案是否為合法的ZIP壓縮檔格式
        /// </summary>
        /// <param name="zipPath">ZIP檔案的完整路徑</param>
        /// <remarks>不會解壓縮檔案，只會嘗試開啟並確認ZIP結構是否正常</remarks>
        /// <returns>如果是合法的ZIP檔案則回傳true，否則回傳false</returns>
        private bool CheckZipFile(string zipPath)
        {
            if (!fileTool.CheckFileExist(zipPath))
                return false;

            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                {
                    return true;
                }
            }
            catch (InvalidDataException ex)
            {
                throw new Exception("該檔案不是合法的 ZIP 格式！");
                throw new Exception($"錯誤訊息：{ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception("發生其他錯誤：");
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 解壓縮指定ZIP檔案中的某個指定檔案至目標資料夾中
        /// </summary>
        /// <param name="zipPath">ZIP壓縮檔的完整路徑</param>
        /// <param name="targetEntryName">欲解壓縮的檔案名稱(ZIP內部路徑)</param>
        /// <param name="outputFilePath">要解壓縮到的完整檔案路徑</param>
        /// <returns>成功解壓指定檔案回傳true，否則回傳false</returns>
        public bool ExtractSingleFileFromZip(string zipPath, string targetEntryName, string outputFilePath)
        {
            if (!CheckZipFile(zipPath))
                return false;

            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                {
                    // 嘗試取得指定檔案的 ZipArchiveEntry(注意：需要完整路徑，區分大小寫)
                    var entry = archive.GetEntry(targetEntryName);
                    if (entry == null)
                    {
                        throw new Exception($"ZIP 中找不到檔案：{targetEntryName}");
                    }

                    // 解壓縮到指定路徑，若檔案已存在則覆寫
                    entry.ExtractToFile(outputFilePath, overwrite: true);
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("解壓縮發生錯誤：");
                throw new Exception(ex.Message);
            }
        }

        public byte[] ExtractSingleFileToMemory(string zipPath, string targetEntryName)
        {
            if (!CheckZipFile(zipPath)) return null;

            using (var zipStream = System.IO.File.OpenRead(zipPath))
            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Read))
            {
                var entry = archive.GetEntry(targetEntryName);
                if (entry == null)
                {
                    throw new Exception($"ZIP 中找不到檔案：{targetEntryName}");
                }

                using (var entryStream = entry.Open())
                using (var ms = new MemoryStream())
                {
                    entryStream.CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }
    }
}
