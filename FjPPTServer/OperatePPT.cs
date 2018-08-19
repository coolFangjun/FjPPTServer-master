using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using System.Windows;
using System.Collections;
namespace PPTDraw.PPTOperate
{
    /// <summary>
    /// PPT文档操作实现类.
    /// </summary>
    public class OperatePPT
    {
        #region=========基本的参数信息=======
        POWERPOINT.ApplicationClass objApp = null;
        POWERPOINT.Presentation objPresSet = null;
        POWERPOINT.SlideShowWindows objSSWs;
        POWERPOINT.SlideShowTransition objSST;
        POWERPOINT.SlideShowSettings objSSS;
        POWERPOINT.SlideRange objSldRng;
        bool bAssistantOn;
        double pixperPoint = 0;
        double offsetx = 0;
        double offsety = 0;
        #endregion
        #region===========操作方法==============
        /// <summary>
        /// 打开PPT文档并播放显示。
        /// </summary>
        /// <param name="filePath">PPT文件路径</param>
        public void PPTOpen(string filePath)
        {
            //防止连续打开多个PPT程序.
            if (this.objApp != null) { return; }
            try
            {
                objApp = new POWERPOINT.ApplicationClass();
                //以非只读方式打开,方便操作结束后保存.
                objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
                //Prevent Office Assistant from displaying alert messages:
                bAssistantOn = objApp.Assistant.On;
                objApp.Assistant.On = false;
                objSSS = this.objPresSet.SlideShowSettings;
                objSSS.Run();
            }
            catch (Exception ex)
            {
                this.objApp.Quit();
            }
        }
      
        /// <summary>
        /// PPT下一页。
        /// </summary>
        public void NextSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Next();
        }
        /// <summary>
        /// PPT上一页。
        /// </summary>
        public void PreviousSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Previous();
        }
      
        /// <summary>
        /// 关闭PPT文档。
        /// </summary>
        public void PPTClose()
        {
            //装备PPT程序。
            if (this.objApp != null)
                this.objApp.Quit();
            GC.Collect();
        }
        #endregion
    }
}