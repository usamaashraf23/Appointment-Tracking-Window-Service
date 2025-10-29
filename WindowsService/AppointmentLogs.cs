using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService
{
    public class AppointmentLogs
    {
        private readonly string _connectionString;
        private readonly string _excelFilePath;
        private readonly string _logFilePath;
        public AppointmentLogs()
        {
            //InitializeComponent();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _connectionString = ConfigurationManager.ConnectionStrings["AppointmentTracking"].ConnectionString;
            _excelFilePath = $@"C:\AppointmentReports\Appointment_Tracking_Report_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx"; 
            _logFilePath = @"C:\AppointmentReports\service_log.txt";
        }

        private void WriteToLog(string message)
        {
            try
            {
                // Ensure directory exists
                var directory = Path.GetDirectoryName(_logFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
            }
            catch
            {
                // If file logging fails, do nothing
            }
        }
        public void GenerateExcelReport()
        {

            WriteToLog("Starting Excel report generation in AppointmentLogs class...");

            // Ensure directory exists
            var directory = Path.GetDirectoryName(_excelFilePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            try
            {
                // Use explicit ExcelPackage constructor
                var fileInfo = new FileInfo(_excelFilePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Appointment Tracking");

                    // Create header row
                    CreateHeader(worksheet);

                    int currentRow = 2;
                    int serialNumber = 1;

                    // Execute all queries and populate data
                    var result = ExecuteQueriesAndPopulateExcel(worksheet, currentRow, serialNumber);
                    currentRow = result.CurrentRow;
                    serialNumber = result.SerialNumber;

                    // REMOVED: AutoFitColumns() - causing the missing method error
                    // Instead, set manual column widths
                    SetManualColumnWidths(worksheet);

                    // Save the Excel file
                    package.Save();
                    WriteToLog($"Excel file saved to: {_excelFilePath}");
                }
            }
            catch(Exception ex) 
            {
                WriteToLog($"Error in GenerateExcelReport: {ex.Message}");
                throw;
            }
        }

        private void SetManualColumnWidths(ExcelWorksheet worksheet)
        {
            // Set reasonable column widths manually
            worksheet.Column(1).Width = 5;  // S#
            worksheet.Column(2).Width = 20; // Events
            worksheet.Column(3).Width = 15; // Total No. of Hits
            worksheet.Column(4).Width = 30; // Success
            worksheet.Column(5).Width = 20; // Exception Reported
            worksheet.Column(6).Width = 10; // Failure
            worksheet.Column(7).Width = 10; // Wrong Hits
            worksheet.Column(8).Width = 25; // Details of Wrong Hits
            worksheet.Column(9).Width = 15; // Remarks
        }

        private void CreateHeader(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "S#";
            worksheet.Cells[1, 2].Value = "Events";
            worksheet.Cells[1, 3].Value = "Total No. of Hits";
            worksheet.Cells[1, 4].Value = "Success";
            worksheet.Cells[1, 5].Value = "Exception Reported";
            worksheet.Cells[1, 6].Value = "Failure";
            worksheet.Cells[1, 7].Value = "Wrong Hits";
            worksheet.Cells[1, 8].Value = "Details of Wrong Hits";
            worksheet.Cells[1, 9].Value = "Remarks";

            // Style the header
            var range = worksheet.Cells[1, 1, 1, 9];
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(171, 231, 178));

            // Add white font color and center alignment
            range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
        }

        // Helper class to return multiple values
        private class PopulateResult
        {
            public int CurrentRow { get; set; }
            public int SerialNumber { get; set; }
        }

        private PopulateResult ExecuteQueriesAndPopulateExcel(ExcelWorksheet worksheet, int startRow, int startSerialNumber)
        {
            int currentRow = startRow;
            int serialNumber = startSerialNumber;

            // 1. Authorize Agent
            var authResult = PopulateAuthorizeAgent(worksheet, currentRow, serialNumber);
            currentRow = authResult.CurrentRow;
            serialNumber = authResult.SerialNumber;

            // 2-4. Patient Verification (multiple sub-queries)
            var patientResult = PopulatePatientVerification(worksheet, currentRow, serialNumber);
            currentRow = patientResult.CurrentRow;
            serialNumber = patientResult.SerialNumber;

            // 5-20. Time Slots (multiple sub-queries)
            var timeSlotsResult = PopulateTimeSlots(worksheet, currentRow, serialNumber);
            currentRow = timeSlotsResult.CurrentRow;
            serialNumber = timeSlotsResult.SerialNumber;

            var addAppointmentResult = PopulateAddAppointment(worksheet, currentRow, serialNumber);
            currentRow = addAppointmentResult.CurrentRow;
            serialNumber = addAppointmentResult.SerialNumber;

            var populateReschedule = PopulateReschedule(worksheet, currentRow, serialNumber);
            currentRow = populateReschedule.CurrentRow;
            serialNumber = populateReschedule.SerialNumber;

            var cancelledAppointments = PopulateCancelledAppointments(worksheet, currentRow, serialNumber);
            currentRow = cancelledAppointments.CurrentRow;
            serialNumber = cancelledAppointments.SerialNumber;

            var labResult = PopulateLabResults(worksheet, currentRow, serialNumber);
            currentRow = labResult.CurrentRow;
            serialNumber = labResult.SerialNumber;

            // 21-23. Task Creation
            var taskResult = PopulateTaskCreation(worksheet, currentRow, serialNumber);
            currentRow = taskResult.CurrentRow;
            serialNumber = taskResult.SerialNumber;

            var prescriptionResult = PopulatePrescription(worksheet, currentRow, serialNumber);
            currentRow = prescriptionResult.CurrentRow;
            serialNumber = prescriptionResult.SerialNumber;


            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateAuthorizeAgent(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            //DebugDateTimeIssues();

            var testQuery = "SELECT COUNT(*) FROM WS_TBL_SMARTTALKPHR_TRACKING WHERE App_source='AI_APPOINTMENT'";
            var totalCount = ExecuteScalarQuery(testQuery);
            WriteToLog($"Total AI_APPOINTMENT records: {totalCount}");

            var query = @"
                SELECT COUNT(*) as HitCount 
                FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
                AND MethodName = 'AuthorizeAgent'";

            var hitCount = ExecuteScalarQuery(query);

            WriteToLog($"Total Hit records: {hitCount}");

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Authorize Agent";
            worksheet.Cells[currentRow, 3].Value = hitCount;
            worksheet.Cells[currentRow, 4].Value = hitCount; // Success count
            worksheet.Cells[currentRow, 5].Value = "No";
            worksheet.Cells[currentRow, 6].Value = 0;

            worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            return new PopulateResult { CurrentRow = currentRow + 1, SerialNumber = serialNumber + 1 };
        }

        //private PopulateResult PopulatePatientVerification(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        //{
        //    // Exact Patient Match
        //    var exactMatchQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GetPatientInformation'";

        //    var exactMatchCount = ExecuteScalarQuery(exactMatchQuery);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "Patient Verification";
        //    worksheet.Cells[currentRow, 3].Value = exactMatchCount;
        //    worksheet.Cells[currentRow, 4].Value = "Exact Patient Match";
        //    worksheet.Cells[currentRow, 5].Value = "No";
        //    worksheet.Cells[currentRow, 6].Value = 0;
        //    currentRow++;
        //    serialNumber++;

        //    // Multiple Patients Match
        //    var multipleMatchQuery = @"
        //        SELECT COUNT(*) as HitCount 
        //        FROM WS_TBL_SMARTTALKPHR_TRACKING 
        //        WHERE App_source='AI_APPOINTMENT'
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GetPatientInformation'";

        //    var multipleMatchCount = ExecuteScalarQuery(multipleMatchQuery);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "";
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
        //        AND CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
        //        AND MethodName = 'GetPatientInformation'";

        //    var patientNotExistsCount = ExecuteScalarQuery(patientNotExistsQuery);

        //    worksheet.Cells[currentRow, 1].Value = serialNumber;
        //    worksheet.Cells[currentRow, 2].Value = "";
        //    worksheet.Cells[currentRow, 3].Value = patientNotExistsCount;
        //    worksheet.Cells[currentRow, 4].Value = "Patient Not Exists";
        //    worksheet.Cells[currentRow, 5].Value = "No";
        //    currentRow++;
        //    serialNumber++;

        //    return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        //}

        private PopulateResult PopulatePatientVerification(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;
            // Get counts for each patient verification type
            var exactMatchQuery = @"SELECT COUNT(*)  
                FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
                AND MethodName = 'GetPatientInformation'
                AND Log_Response LIKE '%""},{""%'";

            var multipleMatchQuery = @"
                SELECT COUNT(*) as HitCount 
                FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
                AND MethodName = 'GetPatientInformation'";

            var patientNotExistsQuery = @"
                SELECT COUNT(*) as HitCount 
                FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND MethodName  IN ('GetPatientInformation','GetPatientDetailsViaName','GetPatientDetailsViaDOB')
                AND  Log_Response LIKE'%""Patient does not exist against this phone number""%'";

            var invalidInputsQuery = @"select COUNT(*) from WS_TBL_SMARTTALKPHR_TRACKING with(nolock,nowait) 
                where App_source='AI_APPOINTMENT' AND TRY_CAST(Exception AS nvarchar)  = 'Exceptions'  
                and MethodName in ('GetPatientInformation','GetPatientDetailsViaName','GetPatientDetailsViaDOB')";

            // Execute queries
            var exactMatchCount = ExecuteScalarQuery(exactMatchQuery);
            var multipleMatchCount = ExecuteScalarQuery(multipleMatchQuery);
            var patientNotExistsCount = ExecuteScalarQuery(patientNotExistsQuery);
            var invalidInputsCount = ExecuteScalarQuery(invalidInputsQuery);

            // Calculate total hits
            var totalHits = exactMatchCount + multipleMatchCount + patientNotExistsCount + invalidInputsCount;

            // Build details string with line breaks
            var details = $"Exact Patient Match: {exactMatchCount}\n" +
                          $"Multiple Patients Match: {multipleMatchCount}\n" +
                          $"Patient Not Exists: {patientNotExistsCount}\n" +
                          $"Invalid Inputs: {invalidInputsCount}";

            // First row - Patient Verification with total hits
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Patient Verification";
            worksheet.Cells[currentRow, 3].Value = totalHits; // Total hits (102)
            worksheet.Cells[currentRow, 4].Value = "Exact Patient Match" +"    " + exactMatchCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (40)
            worksheet.Cells[currentRow, 6].Value = 0;
            worksheet.Cells[currentRow, 7].Value = 0;
            currentRow++;

            // Second row - Multiple Patients Match
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Multiple Patients Match" +"    " + multipleMatchCount;
            worksheet.Cells[currentRow, 5].Value = ""; // Success count for multiple matches
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Third row - Patient Not Exists
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Patient Not Exists" + "    " + patientNotExistsCount;
            worksheet.Cells[currentRow, 5].Value = ""; // Success count for not exists
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Fourth row - Invalid Inputs
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Invalid Inputs" + "    " + invalidInputsCount;
            worksheet.Cells[currentRow, 5].Value = ""; // Success count for invalid inputs
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Merge "Patient Verification" cell across 4 rows
            var patientVerificationCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            patientVerificationCell.Merge = true;
            patientVerificationCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            patientVerificationCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            patientVerificationCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            patientVerificationCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            // Merge "Total Hits" cell across 4 rows and center align
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;
            totalHitsCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            totalHitsCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var patientMatchExceptionCell = worksheet.Cells[startRow, 5, currentRow - 3, 5];
            patientMatchExceptionCell.Merge = true;
            patientMatchExceptionCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            patientMatchExceptionCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var patientVerificationFailure = worksheet.Cells[startRow, 6, currentRow - 3, 6];
            patientVerificationFailure.Merge = true;
            patientVerificationFailure.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            patientVerificationFailure.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateTimeSlots(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            // Get counts for each time slot type
            var forDrHaqQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
                AND MethodName = 'GettimeSlots'
                AND (Log_Request  like '%""ProviderCode"":""100""%')";

            var forAmiPatelQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
                AND MethodName = 'GettimeSlots'
                AND (Log_Request  like '%""ProviderCode"":""55212287""%')";

            var searchFirstAvailableQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING
                WHERE App_source = 'AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
                AND MethodName = 'GettimeSlots'
                AND(Log_Request  like '%""ProviderCode"":""55212287""%')";

            var searchSpecificDateQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING
                WHERE App_source = 'AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
                AND MethodName = 'GettimeSlots'
                AND(Log_Request  like '%""ProviderCode"":""55212287""%')";

            var searchTelehealthSlotQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING
                WHERE App_source = 'AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
                AND MethodName = 'GettimeSlots'
                AND(Log_Request  like '%""ProviderCode"":""55212287""%')";

            // Execute queries
            var forDrHaqCount = ExecuteScalarQuery(forDrHaqQuery);
            var forAmiPatelCount = ExecuteScalarQuery(forAmiPatelQuery);
            var searchFirstAvailableCount = ExecuteScalarQuery(searchFirstAvailableQuery);
            var searchSpecificDateCount = ExecuteScalarQuery(searchSpecificDateQuery);
            var searchTelehealthSlotCount = ExecuteScalarQuery(searchTelehealthSlotQuery);

            // Calculate total hits
            var totalHits = forDrHaqCount + forAmiPatelCount + searchFirstAvailableCount + searchSpecificDateCount + searchTelehealthSlotCount;

            // First row - Time Slots with total hits
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Time Slots";
            worksheet.Cells[currentRow, 3].Value = totalHits; // Total hits (66)
            worksheet.Cells[currentRow, 4].Value = "For Dr. Haq" + "    " + forDrHaqCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            currentRow++;

            // Second row - For Ami Patel
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "For Ami Patel" + "    " + forAmiPatelCount;
            worksheet.Cells[currentRow, 5].Value = 0; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Third row - Search First Available
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Search First Available" + "    " + searchFirstAvailableCount;
            worksheet.Cells[currentRow, 5].Value = ""; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Fourth row - Search Specific Date
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Search Specific Date" + "    " + searchSpecificDateCount;
            worksheet.Cells[currentRow, 5].Value = ""; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Fifth row - Search Telehealth Slot
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Search Telehealth Slot" + "    " + searchTelehealthSlotCount;
            worksheet.Cells[currentRow, 5].Value = ""; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Merge "Time Slots" cell across 5 rows
            var timeSlotsCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            timeSlotsCell.Merge = true;
            timeSlotsCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            timeSlotsCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            timeSlotsCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            timeSlotsCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            // Merge "Total Hits" cell across 5 rows
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;
            totalHitsCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            totalHitsCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var successSlotCell = worksheet.Cells[startRow, 5, currentRow - 1, 5];
            successSlotCell.Merge = true;
            successSlotCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            successSlotCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var failureSlotCell = worksheet.Cells[startRow, 6, currentRow - 1, 6];
            failureSlotCell.Merge = true;
            failureSlotCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            failureSlotCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateAddAppointment(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            // Get counts for each add appointment type
            var appointmentAddedQuery = @"
                SELECT COUNT(*)  FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -34, CONVERT(date, GETDATE()))
                AND MethodName = 'AddNewAppointment'";

            var alreadyScheduledQuery = @"
                SELECT COUNT(*)  FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -34, CONVERT(date, GETDATE()))
                AND MethodName = 'AddNewAppointment'"; 

            var daysMessageQuery = @"
                SELECT COUNT(*)  FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -34, CONVERT(date, GETDATE()))
                AND MethodName = 'AddNewAppointment'";

            var duplicateEntryQuery = @"
                SELECT COUNT(*)  FROM WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -34, CONVERT(date, GETDATE()))
                AND MethodName = 'AddNewAppointment'";

            // Execute queries
            var appointmentAddedCount = ExecuteScalarQuery(appointmentAddedQuery);
            var alreadyScheduledCount = ExecuteScalarQuery(alreadyScheduledQuery);
            var daysMessageCount = ExecuteScalarQuery(daysMessageQuery);
            var duplicateEntryCount = ExecuteScalarQuery(duplicateEntryQuery);

            // Calculate total hits
            var totalHits = appointmentAddedCount + alreadyScheduledCount + daysMessageCount + duplicateEntryCount;

            // First row - Add Appointment with total hits
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Add Appointment";
            worksheet.Cells[currentRow, 3].Value = totalHits; // Total hits (20)
            worksheet.Cells[currentRow, 4].Value = "Appointment Added" + "    " + appointmentAddedCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (15)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            currentRow++;

            // Second row - Already Scheduled Message
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Already Scheduled Message" + "    " + alreadyScheduledCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (0)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Third row - 14 Days Message
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "14 Days Message" + "    " + daysMessageCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (3)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Fourth row - Duplicate Entry
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Duplicate Entry" + "    " + duplicateEntryCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (2)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            currentRow++;

            // Merge "Add Appointment" cell across 4 rows
            var addAppointmentCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            addAppointmentCell.Merge = true;
            addAppointmentCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            addAppointmentCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            addAppointmentCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            addAppointmentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            // Merge "Total Hits" cell across 4 rows
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;
            totalHitsCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            totalHitsCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var exceptionReportedCell = worksheet.Cells[startRow, 5, currentRow - 1, 5];
            exceptionReportedCell.Merge = true;
            exceptionReportedCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            exceptionReportedCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var failureCell = worksheet.Cells[startRow, 6, currentRow - 1, 6];
            failureCell.Merge = true;
            failureCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            failureCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateReschedule(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            // Get counts for each reschedule type
            var hoursRescheduleQuery = @"
                select COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING where App_source='AI_APPOINTMENT'
                and CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
                and MethodName = 'ReschedulePatientAppointmentsAndAppointment'";

            var rescheduledQuery = @"
                select COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING where App_source='AI_APPOINTMENT'
                and CONVERT(date, Logs_Date) = DATEADD(day, -14, CONVERT(date, GETDATE()))
                and MethodName = 'ReschedulePatientAppointmentsAndAppointment'";


            // Execute queries
            var hoursRescheduleCount = ExecuteScalarQuery(hoursRescheduleQuery);
            var rescheduledCount = ExecuteScalarQuery(rescheduledQuery);

            // Calculate total hits
            var totalHits = hoursRescheduleCount + rescheduledCount;

            // First row - Reschedule with total hits
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Reschedule";
            worksheet.Cells[currentRow, 3].Value = totalHits; // Total hits (3)
            worksheet.Cells[currentRow, 4].Value = "24 Hours Reschedule" + "    " + hoursRescheduleCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (1)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 9].Value = "Restrict the Patient to Reschedule due to 24 Hours Check";
            currentRow++;

            // Second row - Rescheduled
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Rescheduled" + "    " + rescheduledCount;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            worksheet.Cells[currentRow, 8].Value = "";
            currentRow++;


            // Merge "Reschedule" cell across 3 rows
            var rescheduleCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            rescheduleCell.Merge = true;
            rescheduleCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            rescheduleCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            rescheduleCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            rescheduleCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            // Merge "Total Hits" cell across 3 rows
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;
            totalHitsCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            totalHitsCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var exceptionReportedCell = worksheet.Cells[startRow, 5, currentRow - 1, 5];
            exceptionReportedCell.Merge = true;
            exceptionReportedCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            exceptionReportedCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            var failureCell = worksheet.Cells[startRow, 6, currentRow - 1, 6];
            failureCell.Merge = true;
            failureCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            failureCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateCancelledAppointments(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            // Get counts for each reschedule type
            var hoursCancelledAppointmentQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -15, CONVERT(date, GETDATE()))
                AND MethodName = 'CancelRestoreAppointments'";

            var CancelledAppointmentQuery = @"
                SELECT COUNT(*)  from WS_TBL_SMARTTALKPHR_TRACKING 
                WHERE App_source='AI_APPOINTMENT'
                AND CONVERT(date, Logs_Date) = DATEADD(day, -15, CONVERT(date, GETDATE()))
                AND MethodName = 'CancelRestoreAppointments'";


            // Execute queries
            var hoursCancelledCount = ExecuteScalarQuery(hoursCancelledAppointmentQuery);
            var cancelledCount = ExecuteScalarQuery(CancelledAppointmentQuery);

            // Calculate total hits
            var totalHits = hoursCancelledCount + cancelledCount;

            // First row - Reschedule with total hits
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Cancelled";
            worksheet.Cells[currentRow, 3].Value = totalHits; // Total hits (3)
            worksheet.Cells[currentRow, 4].Value = "Cancelled" + "    " + cancelledCount;
            worksheet.Cells[currentRow, 5].Value = cancelledCount; // Success count (1)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = "";
            currentRow++;

            // Second row - Rescheduled
            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "24 Hours Cancelled" + "    " + hoursCancelledCount;
            worksheet.Cells[currentRow, 5].Value = cancelledCount; // Success count
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            worksheet.Cells[currentRow, 9].Value = "Restrict the Patient to Cancel due to 24 Hours Check";
            currentRow++;


            // Merge "Reschedule" cell across 3 rows
            var cancelledCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            cancelledCell.Merge = true;
            cancelledCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            cancelledCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            cancelledCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cancelledCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            // Merge "Total Hits" cell across 3 rows
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;
            totalHitsCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            totalHitsCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateLabResults(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            var labResultsQuery = @"
                select Log_Request, Log_Response,Logs_Date  from WS_TBL_SMARTTALKPHR_TRACKING where App_source='AI_APPOINTMENT'
                and CONVERT(date, Logs_Date) = DATEADD(day, -15, CONVERT(date, GETDATE()))
                and MethodName = 'LabResultWithProviderCommentReview'";
            
            var totalLabResults = ExecuteScalarQuery(labResultsQuery);
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Lab Results";
            worksheet.Cells[currentRow, 3].Value = totalLabResults; // Total hits (3)
            worksheet.Cells[currentRow, 4].Value = totalLabResults;
            worksheet.Cells[currentRow, 5].Value = "No"; // Success count (1)
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = "";

            worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            currentRow++;

            serialNumber++;
            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateTaskCreation(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            var taskCreationQuery = @"
                select COUNT(*) from WS_TBL_SMARTTALKPHR_TRACKING where App_source='AI_APPOINTMENT'
                and CONVERT(date, Logs_Date) = DATEADD(day, -1, CONVERT(date, GETDATE()))
                and MethodName = 'SaveUserTask'";

            var taskCreationCount = ExecuteScalarQuery(taskCreationQuery);

            // Add three identical rows for Task Creation
            
                worksheet.Cells[currentRow, 1].Value = serialNumber;
                worksheet.Cells[currentRow, 2].Value = "Task Creation";
                worksheet.Cells[currentRow, 3].Value = taskCreationCount;
                worksheet.Cells[currentRow, 4].Value = taskCreationCount;
                worksheet.Cells[currentRow, 5].Value = "No";
                worksheet.Cells[currentRow, 6].Value = 0;

                worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

                currentRow++;
                serialNumber++;
            

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulatePrescription(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            var prescriptionQuery = @"
                select COUNT(*) from WS_TBL_SMARTTALKPHR_TRACKING where App_source='AI_APPOINTMENT'
                and CONVERT(date, Logs_Date) = DATEADD(day, -60, CONVERT(date, GETDATE()))
                and MethodName = 'GetPrescriptions'";

            var prescriptionCount = ExecuteScalarQuery(prescriptionQuery);

            // Add three identical rows for Task Creation

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Prescription";
            worksheet.Cells[currentRow, 3].Value = prescriptionCount;
            worksheet.Cells[currentRow, 4].Value = prescriptionCount;
            worksheet.Cells[currentRow, 5].Value = "No";
            worksheet.Cells[currentRow, 6].Value = 0;

            worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            currentRow++;

            serialNumber++;


            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private int ExecuteScalarQuery(string query)
        {
            using (var connection = new SqlConnection(_connectionString))
            using (var command = new SqlCommand(query, connection))
            {
                connection.Open();
                var result = command.ExecuteScalar();
                return result != null ? Convert.ToInt32(result) : 0;
            }
        }
    }
}
