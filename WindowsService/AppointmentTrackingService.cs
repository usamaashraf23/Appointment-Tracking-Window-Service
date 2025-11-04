using OfficeOpenXml;
using System;
using System.Data.SqlClient;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace WindowsService
{
    public partial class AppointmentTrackingService : ServiceBase
    {
        private System.Timers.Timer _timer;
        private AppointmentLogs _appointmentLogs;

        public AppointmentTrackingService()
        {
            InitializeComponent();
            _appointmentLogs = new AppointmentLogs();
        }

        protected override void OnStart(string[] args)
        {
            // Remove debugger launch for production
            // System.Diagnostics.Debugger.Launch();

            WriteToLog("=== Service Starting ===");

            // Use thread pool to avoid blocking
            ThreadPool.QueueUserWorkItem(state =>
            {
                InitializeService();
            });
        }

        private void InitializeService()
        {
            try
            {
                WriteToLog("Initializing service components...");

                // Calculate initial delay for 15 minutes
                //DateTime firstRunTime = DateTime.Now.AddMinutes(15);
                DateTime firstRunTime = DateTime.Now;
                double initialDelay = (firstRunTime - DateTime.Now).TotalMilliseconds;

                _timer = new System.Timers.Timer();
                _timer.Interval = initialDelay;
                _timer.Elapsed += OnFirstRun; // ← THIS WAS MISSING!
                _timer.AutoReset = false;
                _timer.Start();

                WriteToLog($"Service fully initialized. First run at: {firstRunTime:HH:mm:ss}");
            }
            catch (Exception ex)
            {
                WriteToLog($"Initialization error: {ex}");
            }
        }

        private void OnFirstRun(object sender, ElapsedEventArgs e)
        {
            try
            {
                WriteToLog("OnFirstRun executed - reconfiguring timer for 24 hours");

                // Remove the first-run handler
                _timer.Elapsed -= OnFirstRun;

                // Set up for 24-hour intervals
                _timer.Interval = 24 * 60 * 60 * 1000; // 24 hours
                _timer.Elapsed += Timer_Elapsed;
                _timer.AutoReset = true;

                // Run immediately
                Timer_Elapsed(null, null);
            }
            catch (Exception ex)
            {
                WriteToLog($"ERROR in OnFirstRun: {ex}");
            }
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                WriteToLog("Starting to generate Excel report...");
                _appointmentLogs.GenerateExcelReport();
                WriteToLog("Excel report generated successfully.");
            }
            catch (Exception ex)
            {
                WriteToLog($"Error generating report: {ex.Message}");
                WriteToLog($"Full error details: {ex}");
            }
        }

        protected override void OnStop()
        {
            WriteToLog("Appointment Tracking Service stopped.");
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Dispose();
            }
        }

        private void WriteToLog(string message)
        {
            string logFilePath = @"C:\AppointmentReports\service_log.txt";
            var directory = Path.GetDirectoryName(logFilePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
        }
    }
}