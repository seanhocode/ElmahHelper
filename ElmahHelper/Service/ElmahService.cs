using ElmahHelper.Model;
using ElmahHelper.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ElmahHelper.Service
{
    public class ElmahService
    {
        private ZipTool zipTool = new ZipTool();
        private FileTool fileTool = new FileTool();

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
    }

}
