using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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
        static UDPServerClass udpServer = new UDPServerClass();
        SynchronizationContext m_SyncContext = null;
        private void Form1_Load(object sender, EventArgs e)
        {
            m_SyncContext = SynchronizationContext.Current;
            //
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

            logText.Text = logText.Text + message+"\n";
            
        }
        void UdpServer_MessageArrived(string Message) {

            WriteLogs("有客服端连接咯");
            try
            {
                Dictionary<string, string> dic = JsonConvert.DeserializeObject<Dictionary<string, string>>(Message);
                string control = dic["control"];
                WriteLogs("收到指令:" + control);
                string info = "";
                Dictionary<string, object> sendDic = new Dictionary<string, object>();
                //组装数据发送给APP
                if (control == "getPPTState")
                {
                    //获取PPT列表
                    sendDic.Add("control", "setPPTState");
                    sendDic.Add("allInfo", GetPPTList());
                    info = "执行获取播放列表";
                    WriteLogs(info);
                }
                if (control == "play")
                {
                    //执行播放指令
                    string name = dic["name"];
                    info = "收到播放"+name+"指令";
                    WriteLogs(info);
                    return;
                }
                if (control == "firstPage")
                {
                    //执行第一页指令
                    info = "收到切换到第一页指令";
                    WriteLogs(info);
                    return;
                }
                if (control == "upPage")
                {
                    //执行上一页指令
                    info = "收到切换到上一页指令";
                    WriteLogs(info);
                    return;
                }
                if (control == "nextPage")
                {
                    //执行下一页指令
                    info = "收到切换到下一页指令";
                    WriteLogs(info);
                    return;
                }
               

                if (sendDic.Count != 0)
                {
                    string str = Newtonsoft.Json.JsonConvert.SerializeObject(sendDic);
                    byte[] data = System.Text.Encoding.UTF8.GetBytes(str);
                    try
                    {
                        udpServer.sendData(data);
                        WriteLogs("发送播放列表成功！");
                    }
                    catch
                    {
                        WriteLogs("回发播放列表失败：" + str );

                    }

                }
                else
                {
                    WriteLogs("播放列表异常");
                }
            }
            catch
            {
                WriteLogs("指令错误！！！！");
            }

        }

        private void OpenPPT() {


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
                    if (aLastName == "ppt"|| aLastName=="pptx") {
                        pptList.Add(str);
                    }
                    
                }
            }
        }
    }
}
