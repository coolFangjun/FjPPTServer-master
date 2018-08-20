using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private void Form1_Load(object sender, EventArgs e)
        {
            m_SyncContext = SynchronizationContext.Current;
            openFileDialog1.Title = "选择要打开的文件";
            openFileDialog1.FileName = "";
            openFileDialog1.InitialDirectory = path+"ppt";   //@是取消转义字符的意思
            openFileDialog1.Filter = "ppt文件|*.ppt;*.pptx;*.pps";
            GetPPTList();
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
                    sendDic.Add("control", "setPPTState");
                    sendDic.Add("allInfo", GetPPTList());
                    info = "执行获取列表";
                    WriteLogs(info);
                }
                if (control == "play")
                {
                    //执行播放指令
                    string name = dic["name"];
                    info = "收到播放"+name+"指令";
                    OpenPPT(path+"ppt/"+ name);
                    return;
                }
                if (control == "firstPage")
                {
                    //执行第一页指令
                    FristAction();
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
                WriteLogs("指令错误！！！！");
            }

        }


        private List<string> GetPPTList() {

            List<string> list = new List<string>();
            GetLists(path+"ppt/",list);
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
            Thread.Sleep(2000);
            ppt.OpenPPT(pptPath);
            WriteLogs("打开PPT" +"'"+ System.IO.Path.GetFileNameWithoutExtension(pptPath) +"'");

        }
        private void FristAction() {

            ppt.fristAction();
            info = ppt.is_open ? "跳转到第一页" : "PPT未打开";
            WriteLogs(info);
        }
        private void NextAction()
        {
            ppt.NextAction();
            info = ppt.is_open ? "跳转到下一页" : "PPT未打开";
            WriteLogs(info);
        }
        private void UpAction()
        {

            ppt.UpAction();
            info = ppt.is_open ? "跳转到上一页" : "PPT未打开";
            WriteLogs(info);
        }
        private void LastAction()
        {

            ppt.LastAction();
            info = ppt.is_open ? "跳转到最后一页" : "PPT未打开";
            WriteLogs(info);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            FristAction();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UpAction();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            NextAction();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            LastAction();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            logText.Text = "";
        }

        private void button7_Click(object sender, EventArgs e)
        {

            ClosePPT(richTextBox1.Text);


        }
        private void ClosePPT(string path = "") {

            Process[] process = Process.GetProcesses();
            foreach (Process prc in process)
            {
                if (prc.ProcessName == "wpp")
                    prc.Kill();
            }


        }
    }
}
