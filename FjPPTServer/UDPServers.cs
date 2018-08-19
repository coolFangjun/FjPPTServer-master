using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.IO;
using System.Diagnostics;

namespace FjPPTServer
{

public class UDPServerClass
{
    public delegate void MessageHandler(string Message);//定义委托事件  
    public event MessageHandler MessageArrived;
    public delegate void ThreadDelegateAction();
    public event ThreadDelegateAction ReloadAction;
    System.Timers.Timer threadTimer = new System.Timers.Timer(4000);
        public UDPServerClass()
    {
        //获取本机可用IP地址  
        IPAddress[] ips = Dns.GetHostAddresses(Dns.GetHostName());
        foreach (IPAddress ipa in ips)
        {
            if (ipa.AddressFamily == AddressFamily.InterNetwork)
            {
                MyIPAddress = ipa;//获取本地IP地址  
                    //Console.WriteLine("本地IP："+ipa);
                break;
            }
        }
        Note_StringBuilder = new StringBuilder();
        PortName = 8888;

    }

    public UdpClient ReceiveUdpClient;

    /// <summary>  
            /// 侦听端口名称  
            /// </summary>  
    public int PortName;

    /// <summary>  
            /// 本地地址  
            /// </summary>  
    public IPEndPoint LocalIPEndPoint;

    /// <summary>  
            /// 日志记录  
            /// </summary>  
    public StringBuilder Note_StringBuilder;
    /// <summary>  
            /// 本地IP地址  
            /// </summary>  
    public IPAddress MyIPAddress;
    public Thread myThread;
        public IPEndPoint client;

    public void Thread_Listen()
    {
        //创建一个线程接收远程主机发来的信息  
        myThread = new Thread(ReceiveData);
        myThread.IsBackground = true;
        myThread.Start();
        threadTimer.AutoReset = true;
        threadTimer.Elapsed += new System.Timers.ElapsedEventHandler(ThreadAction);
        threadTimer.Enabled = true;
    }
        /// <summary>
        /// 监听服务器线程是否结束
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThreadAction(object sender, System.Timers.ElapsedEventArgs e)
        {

            if ((myThread.ThreadState == System.Threading.ThreadState.Stopped||!myThread.IsAlive))
            {

                //线程停止，重启线程
                Debug.WriteLine("监听列表线程已经停止，开始重启");
                ReceiveUdpClient.Close();
                Debug.WriteLine("线程状态1:" + myThread.ThreadState);
                Debug.WriteLine("线程状态2:" + myThread.IsAlive);
                myThread = new Thread(ReceiveData);
                myThread.IsBackground = true;
                myThread.Start();
                //ReloadAction();

                
            }
            Debug.WriteLine("线程状态1:" + myThread.ThreadState);
            Debug.WriteLine("线程状态2:" + myThread.IsAlive);
        }

        /// <summary>  
                /// 接收数据  
                /// </summary>  
        private void ReceiveData()
    {
        IPEndPoint local = new IPEndPoint(MyIPAddress, PortName);
        ReceiveUdpClient = new UdpClient(local);
        IPEndPoint remote = new IPEndPoint(IPAddress.Any, 0);
           
        while (true)
        {
            try
            {
                //关闭udpClient 时此句会产生异常  
                byte[] receiveBytes = ReceiveUdpClient.Receive(ref remote);
                string receiveMessage = Encoding.GetEncoding("utf-8").GetString(receiveBytes, 0, receiveBytes.Length);
                    client = new IPEndPoint(remote.Address,remote.Port);
                    MessageArrived(receiveMessage);

                }
            catch
            {
                break;
            }
        }
    }

        public void sendData(byte[] data) {
            ReceiveUdpClient.Send(data, data.Length, client);
        }
}

}