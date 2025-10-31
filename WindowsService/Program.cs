using System;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace WindowsService
{
    internal static class Program
    {
        static async Task Main()
        {
            //ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[]
            //{
            //    new AppointmentTrackingService()
            //};
            //ServiceBase.Run(ServicesToRun);

            AppointmentLogs appointmentLogs = new AppointmentLogs();
            await appointmentLogs.SendExcelByEmail();

        }
    }
}