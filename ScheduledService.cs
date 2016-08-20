using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Timers;
using System.Configuration;

namespace DispatcherScheduler
{
    public partial class ScheduledService : ServiceBase
    {
        Timer timer = new Timer();
        public ScheduledService()
        {
            InitializeComponent();
        }
      public  static string schedularpath = "";
        protected override void OnStart(string[] args)
        {

            //handle Elapsed event
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);

            //This statement is used to set interval to serviceInterval specified in app.config

            timer.Interval = Convert.ToDouble(ConfigurationManager.AppSettings["serviceInterval"].ToString());

            //enabling the timer
            timer.Enabled = true;
            
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            //TraceService("Another entry at cc " + DateTime.Now);

            string logFilePath = ConfigurationManager.AppSettings["logFilePath"];
            if (!Directory.Exists(logFilePath))
            {
                return;
            }
            schedularpath = logFilePath + "DiapatcherAppLog_" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + "";
            if (!Directory.Exists(schedularpath))
                Directory.CreateDirectory(schedularpath);
            string filepath = schedularpath + @"\" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + "_" + DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second + ".log";
            schedularpath = filepath;
            string JsonConfigPath = ConfigurationManager.AppSettings["JsonConfigPath"];
            if (!File.Exists(JsonConfigPath))
            {
                TraceService("Json file doesnot exist : " + JsonConfigPath);
                return;
            }
            TraceService("service started");//service starts
            CloudCherry objcloud = new CloudCherry();
            objcloud.ImportData();//function to import data into CloudCherry
        }
        protected override void OnStop()
        {
            stopTimer();
        }

        public void stopTimer()
        {
            if (File.Exists(schedularpath))
            TraceService("service stopped");
        }

        //TraceService -- Method used to log activities of the dispatcher app service
        public void TraceService(string content)
        {
            FileStream fs = new FileStream(schedularpath, FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine("(" + DateTime.Now + ") " + content);
            sw.Flush();
            sw.Close();
        }
    }
}
