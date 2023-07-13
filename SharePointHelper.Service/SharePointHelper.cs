using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SharePointHelper.Service
{
    public partial class SharePointHelper : ServiceBase
    {
        private Timer timer;
        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Core.Config.ProcessRules();
            //Core.RecordCenter.ProcessList("http://sp2013dev", "http://sp2013dev", "/", "TestDocs", "Status", "Approved");
        }
        public SharePointHelper()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.timer = new System.Timers.Timer(60000D);  // 30000 milliseconds = 30 seconds
            this.timer.AutoReset = true;
            this.timer.Elapsed += new ElapsedEventHandler(this.timer_Elapsed);
            this.timer.Start();
        }

        protected override void OnStop()
        {
            this.timer.Stop();
        }
    }
}
