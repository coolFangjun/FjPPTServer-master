using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace FjPPTServer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
        Thread myThread;
        Process pr = new Process();//声明一个进程类对象
        static UDPServerClass udpServer = new UDPServerClass();
        PPTControl ppt = new PPTControl();
        SynchronizationContext m_SyncContext = null;
        string info = "";
        [DllImport("user32.dll ")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        //根据任务栏应用程序显示的名称找相应窗口的句柄
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
        private const int SW_RESTORE = 9;
        private void Form1_Load(object sender, EventArgs e)
        {
            m_SyncContext = SynchronizationContext.Current;
            openFileDialog1.Title = "选择要打开的文件";
            openFileDialog1.FileName = "";
            openFileDialog1.InitialDirectory = path+"ppt";   //@是取消转义字符的意思
            openFileDialog1.Filter = "ppt文件|*.ppt;*.pptx;*.pps";
            //GetPPTList();
            myThread = new Thread(StartDUP);
            myThread.Start();

        }
        private void StartDUP()
        {

            //启用socket发送UDP数据
            udpServer.Thread_Listen();
            udpServer.MessageArrived += new UDPServerClass.MessageHandler(UdpServer_MessageArrived);

        }
        private void WriteLogs(string message) {

            m_SyncContext.Post(Logs, message);

        }
        private void Logs(object message) {

            logText.Text += DateTime.Now.ToString()+" "+ message+"\n";
            
        }
        void UdpServer_MessageArrived(string Message) {

            WriteLogs("有客服端连接咯");
            try
            {
                Dictionary<string, string> dic = JsonConvert.DeserializeObject<Dictionary<string, string>>(Message);
                string control = dic["control"];
                WriteLogs("收到指令:" + control); 
                Dictionary<string, object> sendDic = new Dictionary<string, object>();
                //组装数据发送给APP
                if (control == "getPPTState")
                {
                    //获取PPT列表
                    string type = dic["type"];
                    string str = type == "0" ? path + "ppt/乡镇规划/" : path + "ppt/招商引资/";
                    sendDic.Add("control", "setPPTState");
                    sendDic.Add("allInfo", GetPPTList(str));
                    info = "执行获取列表";
                    WriteLogs(info);
                }
                if (control == "play")
                {
                    //执行播放指令
                    string name = dic["name"];
                    string type = dic["type"];
                    string str = type == "0" ? path + "ppt/乡镇规划/" : path + "ppt/招商引资/";
                    info = "收到播放"+name+"指令";
                    OpenPPT(str+name);
                    return;
                }
                if (control == "fullScreen")
                {
                    //执行全屏
                    ScreenAction();
                    return;
                }
                if (control == "upPage")
                {
                    //执行上一页指令
                    UpAction();
                    return;
                }
                if (control == "nextPage")
                {
                    //执行下一页指令
                    NextAction();
                    return;
                }
                if (control == "close") {
                    ClosePPT();
                    return;
                }
                if (control == "esc")
                {
                    EscPPT();
                    return;
                }

                if (sendDic.Count != 0)
                {
                    string str = Newtonsoft.Json.JsonConvert.SerializeObject(sendDic);
                    byte[] data = System.Text.Encoding.UTF8.GetBytes(str);
                    try
                    {
                        udpServer.sendData(data);
                        WriteLogs("发送列表成功！");
                    }
                    catch
                    {
                        WriteLogs("回发列表失败：" + str );

                    }

                }
                else
                {
                    WriteLogs("列表异常");
                }
            }
            catch
            {
                WriteLogs("请先执行打开PPT！！");
            }

        }

        private void ActiveAction() {

            string pName = "POWERPNT";//要启动的进程名称，可以在任务管理器里查看，一般是不带.exe后缀的;
            Process[] temp = Process.GetProcessesByName(pName);//在所有已启动的进程中查找需要的进程；
            if (temp.Length >0)//如果查找到
            {
                for (int i = 0; i < temp.Length; i++)
                {

                    IntPtr handle = temp[i].MainWindowHandle;
                    SwitchToThisWindow(handle, true);    // 激活，显示在最前
                }
            }
            else
            {
                Process.Start(pName + ".exe");//否则启动进程
            }


        }


        private List<string> GetPPTList(string str = "") {

            List<string> list = new List<string>();
            GetLists(str.Length==0? path + "ppt/乡镇规划/":str, list);
            return list;

        }
        public void GetLists(string dir, List<string> pptList)
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹  
                {
                    GetLists(fsinfo.FullName, pptList);//递归调用  

                }
                else
                {
                    // Console.WriteLine(fsinfo.FullName);//输出文件的全部路径  
                    string str = fsinfo.Name;
                    string aLastName = str.Substring(str.LastIndexOf(".") + 1, (str.Length - str.LastIndexOf(".") - 1)); //扩展名
                    if (aLastName == "ppt"|| aLastName=="pptx"|| aLastName == "pps") {
                        pptList.Add(str);
                        //|| fsinfo.Attributes != FileAttributes.Hidden
                    }
                    
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Text = openFileDialog1.FileName;
                OpenPPT(richTextBox1.Text);

            }
        }

        private void OpenPPT(string pptPath)
        {
            pr.StartInfo.FileName = pptPath;
            pr.Start();
            ScreenAction();
            WriteLogs("打开PPT" +"'"+ System.IO.Path.GetFileNameWithoutExtension(pptPath) +"'");

        }

        private void ScreenAction()
        {
            ActiveAction();
            Thread.Sleep(2000);
            SendKeys.SendWait("{F5}");
            info = "全屏播放";
            WriteLogs(info);
        }
        private void NextAction()
        {
            //ppt.NextAction();
            SendKeys.SendWait("{PGDN}");
            info = "跳转到下一页";
            WriteLogs(info);
        }
        private void UpAction()
        {

            //ppt.UpAction();
            SendKeys.SendWait("{PGUP}");
            info = "跳转到上一页";
            WriteLogs(info);
        }
        private void EscPPT()
        {
            //ppt.LastAction();
            SendKeys.SendWait("{ESC}");
            info ="退出全屏";
            WriteLogs(info);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            logText.Text = "";
        }

        private void button7_Click(object sender, EventArgs e)
        {

            ClosePPT();
        }
        
        private void ClosePPT() {
            info = "关闭PPT";
            WriteLogs(info);
            Process[] process = Process.GetProcesses();
            foreach (Process prc in process)
            {
                if (prc.ProcessName == "wpp"|| prc.ProcessName == "POWERPNT")
                    prc.Kill();
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            ScreenAction();
        }
    }
}
