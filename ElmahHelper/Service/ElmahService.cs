using ElmahHelper.Model;
using ElmahHelper.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ElmahHelper.Service
{
    public class ElmahService
    {
        private ZipTool zipTool = new ZipTool();
        private FileTool fileTool = new FileTool();
        private FormControlTool controlTool = new FormControlTool();

        /// <summary>
        /// 取得Elmah清單
        /// </summary>
        /// <param name="elmahFolderPath"></param>
        /// <returns></returns>
        public IList<Elmah> GetElmahList(string elmahFolderPath, DateTime? startTime, DateTime? endTime, string fileNameContain = "", string messageContain = "", string detailContain = "")
        {
            IList<Elmah> elmahList = new List<Elmah>();

            IList<string> filePathList = fileTool.GetAllFileInFolder(elmahFolderPath);

            startTime = startTime == null ? new DateTime(1900, 1, 1) : startTime;
            endTime = endTime == null ? new DateTime(2099, 12, 31) : endTime;

            DateTime elmahTime;

            if(filePathList != null)
                foreach (string filePath in filePathList)
                {
                    if(filePath.EndsWith(".zip")){
                        foreach (string fileName in zipTool.GetFileNameInZip(filePath)){
                            elmahTime = Elmah.GetElmahFileNameData(fileName).Value.ElmahTime;
                            if ((elmahTime >= startTime) && (elmahTime <= endTime))
                                elmahList.Add(new Elmah(fileName, filePath));
                        }
                    }
                    else{
                        elmahTime = Elmah.GetElmahFileNameData(Path.GetFileName(filePath)).Value.ElmahTime;
                        if ((elmahTime >= startTime) && (elmahTime <= endTime))
                            elmahList.Add(new Elmah(filePath));
                    }
                }

            elmahList = elmahList
                            .Where(
                                elmah => elmah.FileName.Contains(fileNameContain)
                                && elmah.ElmahError.Message.Contains(messageContain)
                                && elmah.ElmahError.GetDetail().Contains(detailContain))
                            .ToList();

            return elmahList.OrderByDescending(elmah => elmah.ElmahError.Time).ToList();
        }

        /// <summary>
        /// 刪除Elmah
        /// </summary>
        /// <remarks>如檔案在zip裡，會刪除整個zip。會備份至BackUp\yyyyMMdd-HHmmss</remarks>
        /// <param name="gridErrorList">欲刪除的ErrorList</param>
        /// <param name="elmahList">SourceElmah清單</param>
        public void DeleteElmah(IList<Error> gridErrorList, IList<Elmah> elmahList)
        {
            string filePathMsg = string.Empty
                 , backupFolderPath = Path.Combine(FileTool.CurrentFolder, "BackUp", $"{DateTime.Now.ToString("yyyyMMdd-HHmmss")}");

            //Step.1 取得Grid上的Elmah
            IList<Elmah> deleteElmahList = elmahList.Where(elmah => gridErrorList.Any(error => error.ErrorID.Contains(elmah.GUID))).ToList();

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
            }
            else
            {
                MessageBox.Show("無要刪除的檔案");
            }
        }
    }

}
