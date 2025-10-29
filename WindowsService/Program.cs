using System;
using System.ServiceProcess;

namespace WindowsService
{
    internal static class Program
    {
        static void Main()
        {
            //ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[]
            //{
            //    new AppointmentTrackingService()
            //};
            //ServiceBase.Run(ServicesToRun);

            AppointmentLogs appointmentLogs = new AppointmentLogs();
            appointmentLogs.GenerateExcelReport();

        }
    }
}