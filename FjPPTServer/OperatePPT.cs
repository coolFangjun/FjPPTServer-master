using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
// 添加PowerPoint引用
using PPt = Microsoft.Office.Interop.PowerPoint;

namespace FjPPTServer
{
    // 自动遥控幻灯片
    // 这里需要注意的是，在普通视图和阅读模式中实现切换幻灯片的操作不一样
    // 参考信息有：http://msdn.microsoft.com/en-us/library/bb251394(v=office.12).aspx
    //http://support.microsoft.com/kb/316126 
    // http://www.codeproject.com/Questions/73414/Getting-Running-Instance-of-Powerpoint-in-C
    public partial class PPTControl 
    {
        // 定义PowerPoint应用程序对象
        PPt.Application pptApplication;
        // 定义演示文稿对象
        PPt.Presentation presentation;
        // 定义幻灯片集合对象
        PPt.Slides slides;
        // 定义单个幻灯片对象
        PPt.Slide slide;

        // 幻灯片的数量
        int slidescount;
        // 幻灯片的索引
        int slideIndex;
        public bool is_open = false;
        public void OpenPPT(string pptPath)
        {
            // 必须先运行幻灯片，下面才能获得PowerPoint应用程序，否则会出现异常
            // 获得正在运行的PowerPoint应用程序
            try
            {
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

            }
            catch
            {
                MessageBox.Show("启动幻灯片失败", "Error", MessageBoxButtons.OKCancel);
            }
            if (pptApplication != null)
            {

               
                try
                {

                    // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                    // 然而在阅读模式下，这种方式会出现异常
                    // 获得当前选中的幻灯片
                    is_open = true;
                    Thread.Sleep(1000);
                    //获得演示文稿对象
                    presentation = pptApplication.ActivePresentation;
                    // 获得幻灯片对象集合
                    slides = presentation.Slides;
                    // 获得幻灯片的数量
                    slidescount = slides.Count;
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                   
                }
                catch
                {
                    slide = slide ?? pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        // 第一页事件
        public void fristAction()
        {
            if (!is_open) {

                return;
            }    
            try
            {
                // 在普通视图中调用Select方法来选中第一张幻灯片
                slides[1].Select();         
                slide = slides[1];
            }
            catch
            {
                // 在阅读模式下使用下面的方式来切换到第一张幻灯片
                pptApplication.SlideShowWindows[1].View.First();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }

        // 最后一页
        public void LastAction()
        {
            if (!is_open)
            {
                return;
            }
            try
            {
                slides[slidescount].Select();
                slide = slides[slidescount];
            }
            catch
            {
                // 在阅读模式下使用下面的方式来切换到最后幻灯片
                pptApplication.SlideShowWindows[1].View.Last();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }

        // 切换到下一页幻灯片
        public void NextAction()
        {
            if (!is_open)
            {
                return;
            }
            slideIndex = slide.SlideIndex + 1;      
            if (slideIndex > slidescount)
            {
                Debug.WriteLine("已经是最后一页了");
                fristAction();
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到下一张幻灯片
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        // 切换到上一页幻灯片
        public void UpAction()
        {
            if (!is_open)
            {

                return;
            }
            slideIndex = slide.SlideIndex - 1;   
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到上一张幻灯片
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                Debug.WriteLine("已经是第一页了");
            }
        }
    }
}
