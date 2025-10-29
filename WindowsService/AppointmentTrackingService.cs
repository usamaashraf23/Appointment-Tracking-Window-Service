using System;
using System.Data.SqlClient;
using System.IO;
using System.ServiceProcess;
using System.Timers;
using OfficeOpenXml;

namespace WindowsService
{
    public partial class AppointmentTrackingService : ServiceBase
    {
        private Timer _timer;
        
        private AppointmentLogs _appointmentLogs;

        public AppointmentTrackingService()
        {
            InitializeComponent();
            _appointmentLogs = new AppointmentLogs();
        }

        protected override void OnStart(string[] args)
        {
            System.Diagnostics.Debugger.Launch();

            WriteToLog("Appointment Tracking Service started.");

            // Set up timer to run every 24 hours
            _timer = new Timer();
            _timer.Interval = 24 * 60 * 60 * 1000; // 24 hours in milliseconds
            _timer.Elapsed += Timer_Elapsed;
            _timer.Start();

            // Run immediately on start (optional)
            Timer_Elapsed(null, null);
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

        //private void GenerateExcelReport()
        //{
        //    // Ensure directory exists
        //    var directory = Path.GetDirectoryName(_excelFilePath);
        //    if (!Directory.Exists(directory))
        //    {
        //        Directory.CreateDirectory(directory);
        //    }

        //    // Use explicit ExcelPackage constructor
        //    var fileInfo = new FileInfo(_excelFilePath);
        //    using (var package = new ExcelPackage(fileInfo))
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("Appointment Tracking");

        //        // Create header row
        //        CreateHeader(worksheet);

        //        int currentRow = 2;
        //        int serialNumber = 1;

        //        // Execute all queries and populate data
        //        var result = ExecuteQueriesAndPopulateExcel(worksheet, currentRow, serialNumber);
        //        currentRow = result.CurrentRow;
        //        serialNumber = result.SerialNumber;

        //        // REMOVED: AutoFitColumns() - causing the missing method error
        //        // Instead, set manual column widths
        //        SetManualColumnWidths(worksheet);

        //        // Save the Excel file
        //        package.Save();
        //        WriteToLog($"Excel file saved to: {_excelFilePath}");
        //    }
        //}

        //private void SetManualColumnWidths(ExcelWorksheet worksheet)
        //{
        //    // Set reasonable column widths manually
        //    worksheet.Column(1).Width = 5;  // S#
        //    worksheet.Column(2).Width = 20; // Events
        //    worksheet.Column(3).Width = 15; // Total No. of Hits
        //    worksheet.Column(4).Width = 20; // Success
        //    worksheet.Column(5).Width = 20; // Exception Reported
        //    worksheet.Column(6).Width = 10; // Failure
        //    worksheet.Column(7).Width = 10; // Wrong Hits
        //    worksheet.Column(8).Width = 25; // Details of Wrong Hits
        //    worksheet.Column(9).Width = 15; // Remarks
        //}

        //private void CreateHeader(ExcelWorksheet worksheet)
        //{
        //    worksheet.Cells[1, 1].Value = "S#";
        //    worksheet.Cells[1, 2].Value = "Events";
        //    worksheet.Cells[1, 3].Value = "Total No. of Hits";
        //    worksheet.Cells[1, 4].Value = "Success";
        //    worksheet.Cells[1, 5].Value = "Exception Reported";
        //    worksheet.Cells[1, 6].Value = "Failure";
        //    worksheet.Cells[1, 7].Value = "Wrong Hits";
        //    worksheet.Cells[1, 8].Value = "Details of Wrong Hits";
        //    worksheet.Cells[1, 9].Value = "Remarks";

        //    // Style the header
        //    var range = worksheet.Cells[1, 1, 1, 9];
        //    range.Style.Font.Bold = true;
        //    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;

        //    // Use RGB values for background color
        //    range.Style.Fill.BackgroundColor.SetColor(255, 192, 192, 192); // Solid LightGray
        //}

        //// Helper class to return multiple values
        //private class PopulateResult
        //{
        //    public int CurrentRow { get; set; }
        //    public int SerialNumber { get; set; }
        //}

        //private PopulateResult ExecuteQueriesAndPopulateExcel(ExcelWorksheet worksheet, int startRow, int startSerialNumber)
        //{
        //    int currentRow = startRow;
        //    int serialNumber = startSerialNumber;

        //    // 1. Authorize Agent
        //    var authResult = PopulateAuthorizeAgent(worksheet, currentRow, serialNumber);
        //    currentRow = authResult.CurrentRow;
        //    serialNumber = authResult.SerialNumber;

        //    // 2-4. Patient Verification (multiple sub-queries)
        //    var patientResult = PopulatePatientVerification(worksheet, currentRow, serialNumber);
        //    currentRow = patientResult.CurrentRow;
        //    serialNumber = patientResult.SerialNumber;

        //    // 5-20. Time Slots (multiple sub-queries)
        //    var timeSlotsResult = PopulateTimeSlots(worksheet, currentRow, serialNumber);
        //    currentRow = timeSlotsResult.CurrentRow;
        //    serialNumber = timeSlotsResult.SerialNumber;

        //    // 21-23. Task Creation
        //    var taskResult = PopulateTaskCreation(worksheet, currentRow, serialNumber);
        //    currentRow = taskResult.CurrentRow;
        //    serialNumber = taskResult.SerialNumber;

        //    return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        //}

        //private PopulateResult PopulateAuthorizeAgent(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        //{
        //    var query = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
        //        AND MethodName = 'AuthorizeAgent'
        //        AND Log_Request NOT LIKE '%""practice_code"":""1011163""%'";

        //    var hitCount = ExecuteScalarQuery(query);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "Authorize Agent";
        //    worksheet.Cells[currentRow, 3].Value = hitCount;
        //    worksheet.Cells[currentRow, 4].Value = hitCount; // Success count
        //    worksheet.Cells[currentRow, 5].Value = "No";
        //    worksheet.Cells[currentRow, 6].Value = 0;

        //    return new PopulateResult { CurrentRow = currentRow + 1, SerialNumber = serialNumber + 1 };
        //}

        //private PopulateResult PopulatePatientVerification(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        //{
        //    // Exact Patient Match
        //    var exactMatchQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GetPatientInformation'
        //        AND Log_Request NOT LIKE '%""PracticeCode"":1011163%'
        //        AND Log_Response NOT LIKE '%""},{""%'";

        //    var exactMatchCount = ExecuteScalarQuery(exactMatchQuery);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "Patient Verification";
        //    worksheet.Cells[currentRow, 3].Value = exactMatchCount;
        //    worksheet.Cells[currentRow, 4].Value = "Asset Patient Match";
        //    worksheet.Cells[currentRow, 5].Value = "No";
        //    worksheet.Cells[currentRow, 6].Value = 0;
        //    currentRow++;
        //    serialNumber++;

        //    // Multiple Patients Match
        //    var multipleMatchQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GetPatientInformation'
        //        AND Log_Request NOT LIKE '%""PracticeCode"":1011163%'
        //        AND Log_Response LIKE '%""},{""%'";

        //    var multipleMatchCount = ExecuteScalarQuery(multipleMatchQuery);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "Patient Verification";
        //    worksheet.Cells[currentRow, 3].Value = multipleMatchCount;
        //    worksheet.Cells[currentRow, 4].Value = "Multiple Patients Match";
        //    worksheet.Cells[currentRow, 5].Value = "No";
        //    currentRow++;
        //    serialNumber++;

        //    // Patient Not Exists
        //    var patientNotExistsQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GetPatientInformation'
        //        AND Log_Request NOT LIKE '%""PracticeCode"":1011163%'
        //        AND (Log_Response LIKE '%Patient not found%' OR Log_Response LIKE '%No patient found%')";

        //    var patientNotExistsCount = ExecuteScalarQuery(patientNotExistsQuery);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "Patient Verification";
        //    worksheet.Cells[currentRow, 3].Value = patientNotExistsCount;
        //    worksheet.Cells[currentRow, 4].Value = "Patient Not Exists";
        //    worksheet.Cells[currentRow, 5].Value = "No";
        //    currentRow++;
        //    serialNumber++;

        //    return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        //}

        //private PopulateResult PopulateTimeSlots(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        //{
        //    // Total Time Slots Hits
        //    var totalTimeSlotsQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -3, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GettimeSlots'
        //        AND Log_Request NOT LIKE '%""PracticeCode"":""1011163""%'
        //        AND (Log_Request LIKE '%""ProviderCode"":""55212287""%' OR Log_Request LIKE '%""ProviderCode"":""100""%')";

        //    var totalTimeSlotsCount = ExecuteScalarQuery(totalTimeSlotsQuery);

        //    // Add multiple rows for Time Slots with same total count but different success types
        //    string[] successTypes = {
        //        "Invalid Inputs", "For Dr. Liga", "For Ami Patel", "Search First Available",
        //        "Search Specific Date", "Search Telehealth Slot", "Appointment Added",
        //        "All Ready Scheduled Message", "A Days Message", "Duplicate Entry",
        //        "24 Hours Reschedule", "Rescheduled", "Cancelled", "24 Hours Cancellation"
        //    };

        //    foreach (var successType in successTypes)
        //    {
        //        worksheet.Cells[currentRow, 1].Value = serialNumber;
        //        worksheet.Cells[currentRow, 2].Value = "Time Slots";
        //        worksheet.Cells[currentRow, 3].Value = totalTimeSlotsCount;
        //        worksheet.Cells[currentRow, 4].Value = successType;
        //        worksheet.Cells[currentRow, 5].Value = "No";
        //        currentRow++;
        //        serialNumber++;
        //    }

        //    return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        //}

        //private PopulateResult PopulateTaskCreation(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        //{
        //    var taskCreationQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
        //        AND MethodName = 'SaveUserTask'
        //        AND Log_Request NOT LIKE '%""practiceCode"":""1011163""%'";

        //    var taskCreationCount = ExecuteScalarQuery(taskCreationQuery);

        //    // Add three identical rows for Task Creation
        //    for (int i = 0; i < 3; i++)
        //    {
        //        worksheet.Cells[currentRow, 1].Value = serialNumber;
        //        worksheet.Cells[currentRow, 2].Value = i == 0 ? "Task Creation" : "Tasks Creation";
        //        worksheet.Cells[currentRow, 3].Value = taskCreationCount;
        //        worksheet.Cells[currentRow, 4].Value = taskCreationCount;
        //        worksheet.Cells[currentRow, 5].Value = "No";
        //        worksheet.Cells[currentRow, 6].Value = 0;
        //        currentRow++;
        //        serialNumber++;
        //    }

        //    return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        //}

        //private int ExecuteScalarQuery(string query)
        //{
        //    using (var connection = new SqlConnection(_connectionString))
        //    using (var command = new SqlCommand(query, connection))
        //    {
        //        connection.Open();
        //        var result = command.ExecuteScalar();
        //        return result != null ? Convert.ToInt32(result) : 0;
        //    }
        //}
    }
}