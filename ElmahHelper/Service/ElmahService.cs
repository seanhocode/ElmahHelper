using ElmahHelper.Model;
using ElmahHelper.Tools;
using System;
using System.Collections.Generic;
using System.IO;

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
        public IList<Elmah> GetElmahList(string elmahFolderPath, DateTime startTime, DateTime endTime)
        {
            IList<Elmah> elmahList = new List<Elmah>();

            IList<string> filePathList = fileTool.GetAllFileInFolder(elmahFolderPath);

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

            return elmahList;
        }
    }

}
