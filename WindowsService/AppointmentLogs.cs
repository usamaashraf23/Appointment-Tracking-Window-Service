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
        private readonly EmailService _emailService;
        public AppointmentLogs()
        {
            //InitializeComponent();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _connectionString = ConfigurationManager.ConnectionStrings["AppointmentTracking"].ConnectionString;
            _excelFilePath = $@"C:\AppointmentReports\FDA Agent API Hits Report {DateTime.Now.AddDays(-1):dd-MM-yyyy hh-mm-ss}.xlsx";
            _logFilePath = @"C:\AppointmentReports\service_log.txt";
            _emailService = new EmailService();
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
                var fileInfo = new FileInfo(_excelFilePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Appointment Tracking");

                    CreateHeader(worksheet);

                    int currentRow = 2;
                    int serialNumber = 1;

                    var result = ExecuteQueriesAndPopulateExcel(worksheet, currentRow, serialNumber);
                    currentRow = result.CurrentRow;
                    serialNumber = result.SerialNumber;

                    SetManualColumnWidths(worksheet);

                    var dataRange = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column];
                    dataRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
                    dataRange.Style.Border.Top.Style = dataRange.Style.Border.Bottom.Style = dataRange.Style.Border.Left.Style = dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    var columnCellRange = worksheet.Cells[2, 6, worksheet.Dimension.End.Row, 7];
                    columnCellRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    var columnCellRange2 = worksheet.Cells[2, 2, worksheet.Dimension.End.Row, 3];
                    columnCellRange2.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    columnCellRange2.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

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
            worksheet.Column(1).Width = 5;   // S#
            worksheet.Column(2).Width = 20;  // Events
            worksheet.Column(3).Width = 15;  // Total No. of Hits
            worksheet.Column(4).Width = 30;  // Success Description (first sub-column)
            worksheet.Column(5).Width = 10;  // Success Count (second sub-column)
            worksheet.Column(6).Width = 20;  // Exception Reported
            worksheet.Column(7).Width = 10;  // Failure
            worksheet.Column(8).Width = 10;  // Wrong Hits
            worksheet.Column(9).Width = 25;  // Details of Wrong Hits
            worksheet.Column(10).Width = 60; // Remarks
        }

        private void CreateHeader(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "S#";
            worksheet.Cells[1, 2].Value = "Events";
            worksheet.Cells[1, 3].Value = "Total No. of Hits";
            worksheet.Cells[1, 4].Value = "Success";
            worksheet.Cells[1, 4, 1, 5].Merge = true;
            worksheet.Cells[1, 6].Value = "Exception Reported";
            worksheet.Cells[1, 7].Value = "Failure";
            worksheet.Cells[1, 8].Value = "Wrong Hits";
            worksheet.Cells[1, 9].Value = "Details of Wrong Hits";
            worksheet.Cells[1, 10].Value = "Remarks";

            // Style the header
            var range = worksheet.Cells[1, 1, 1, 10];
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(171, 231, 178));

            // Add white font color and center alignment
            range.Style.Font.Color.SetColor(System.Drawing.Color.Black);
            range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
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
            var query = @"EXEC sp_Appointment_Tracking_log @MethodName = 'AuthorizeAgent'";
            var hitCount = ExecuteScalarQuery(query);

            var authorizeAgentExceptionsQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'AuthorizeAgent'";
            var authorizeAgentExceptionsCount = ExecuteScalarQuery(authorizeAgentExceptionsQuery);

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Authorize Agent";
            worksheet.Cells[currentRow, 3].Value = hitCount;

            // Merge success description and count for single row items
            worksheet.Cells[currentRow, 4].Value = hitCount;
            worksheet.Cells[currentRow, 4, currentRow, 5].Merge = true;

            worksheet.Cells[currentRow, 6].Value = authorizeAgentExceptionsCount;
            worksheet.Cells[currentRow, 7].Value = 0;

            worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            return new PopulateResult { CurrentRow = currentRow + 1, SerialNumber = serialNumber + 1 };
        }

        private PopulateResult PopulatePatientVerification(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            var exactMatchQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'ExactPatientMatch'";

            var multipleMatchQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'MultiplePatientsMatch'";

            var patientNotExistsQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'PatientNotExists'";

            var invalidInputsQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'InvalidInputs'";

            var exactMatchExceptionsQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'ExactPatientMatch'";

            var multipleMatchExceptionsQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'MultiplePatientsMatch'";


            var exactMatchCount = ExecuteScalarQuery(exactMatchQuery);
            var multipleMatchCount = ExecuteScalarQuery(multipleMatchQuery);
            var patientNotExistsCount = ExecuteScalarQuery(patientNotExistsQuery);
            var invalidInputsCount = ExecuteScalarQuery(invalidInputsQuery);
            //var exactMatchExceptionsCount = ExecuteScalarQuery(exactMatchExceptionsQuery);
            //var multipleMatchExceptionsCount = ExecuteScalarQuery(multipleMatchExceptionsQuery);

            var totalHits = exactMatchCount + multipleMatchCount + patientNotExistsCount + invalidInputsCount;
            var totalExceptions = /*exactMatchExceptionsCount + multipleMatchExceptionsCount;*/ 0;
            var details = $"Exact Patient Match: {exactMatchCount}\n" +
                          $"Multiple Patients Match: {multipleMatchCount}\n" +
                          $"Patient Not Exists: {patientNotExistsCount}\n" +
                          $"Invalid Inputs: {invalidInputsCount}";

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Patient Verification";
            worksheet.Cells[currentRow, 3].Value = totalHits;
            worksheet.Cells[currentRow, 4].Value = "Exact Patient Match";
            worksheet.Cells[currentRow, 5].Value = exactMatchCount;
            worksheet.Cells[currentRow, 6].Value = totalExceptions;
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Multiple Patients Match";
            worksheet.Cells[currentRow, 5].Value = multipleMatchCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Patient Not Exists";
            worksheet.Cells[currentRow, 5].Value = patientNotExistsCount;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Invalid Inputs";
            worksheet.Cells[currentRow, 5].Value = invalidInputsCount;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            var patientVerificationCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            patientVerificationCell.Merge = true;
            patientVerificationCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            patientVerificationCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            // Merge "Total Hits" cell across 4 rows and center align
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;

            var patientMatchExceptionCell = worksheet.Cells[startRow, 6, currentRow - 3, 6];
            patientMatchExceptionCell.Merge = true;

            var patientVerificationFailure = worksheet.Cells[startRow, 7, currentRow - 3, 7];
            patientVerificationFailure.Merge = true;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateTimeSlots(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            var forDrHaqQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'DrHaqTimeSlot'";

            var forAmiPatelQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'AmiPatelTimeSlot'";

            var searchFirstAvailableQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'FirstAvailableSlot'";

            var searchSpecificDateQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'SpecificDateSlot'";

            var searchTelehealthSlotQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'TelehealthSlot'";

            var forDrHaqExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'DrHaqTimeSlot'";
            var forAmiPatelExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'AmiPatelTimeSlot'";
            var searchFirstAvailableExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'FirstAvailableSlot'";
            var searchSpecificDateExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'SpecificDateSlot'";
            var searchTelehealthSlotExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'TelehealthSlot'";

            var forDrHaqCount = ExecuteScalarQuery(forDrHaqQuery);
            var forAmiPatelCount = ExecuteScalarQuery(forAmiPatelQuery);
            var searchFirstAvailableCount = ExecuteScalarQuery(searchFirstAvailableQuery);
            var searchSpecificDateCount = ExecuteScalarQuery(searchSpecificDateQuery);
            var searchTelehealthSlotCount = ExecuteScalarQuery(searchTelehealthSlotQuery);

            var forDrHaqExceptionCount = ExecuteScalarQuery(forDrHaqExceptionQuery);
            var forAmiPatelExceptionCount = ExecuteScalarQuery(forAmiPatelExceptionQuery);
            var searchFirstAvailableExceptionCount = ExecuteScalarQuery(searchFirstAvailableExceptionQuery);
            var searchSpecificDateExceptionCount = ExecuteScalarQuery(searchSpecificDateExceptionQuery);
            var searchTelehealthSlotExceptionCount = ExecuteScalarQuery(searchTelehealthSlotExceptionQuery);

            var totalHits = forDrHaqCount + forAmiPatelCount + searchFirstAvailableCount + searchSpecificDateCount + searchTelehealthSlotCount;
            var timeSlotTotalExceptionCount = forDrHaqExceptionCount + forAmiPatelExceptionCount + searchFirstAvailableExceptionCount + searchSpecificDateExceptionCount + searchTelehealthSlotExceptionCount;;

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Time Slots";
            worksheet.Cells[currentRow, 3].Value = totalHits;
            worksheet.Cells[currentRow, 4].Value = "For Dr. Haq";
            worksheet.Cells[currentRow, 5].Value = forDrHaqCount;
            worksheet.Cells[currentRow, 6].Value = timeSlotTotalExceptionCount;
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "For Ami Patel";
            worksheet.Cells[currentRow, 5].Value = forAmiPatelCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Search First Available";
            worksheet.Cells[currentRow, 5].Value = searchFirstAvailableCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Search Specific Date";
            worksheet.Cells[currentRow, 5].Value = searchSpecificDateCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Search Telehealth Slot";
            worksheet.Cells[currentRow, 5].Value = searchTelehealthSlotCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            var timeSlotsCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            timeSlotsCell.Merge = true;
            timeSlotsCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            timeSlotsCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;

            var successSlotCell = worksheet.Cells[startRow, 6, currentRow - 1, 6];
            successSlotCell.Merge = true;

            var failureSlotCell = worksheet.Cells[startRow, 7, currentRow - 1, 7];
            failureSlotCell.Merge = true;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateAddAppointment(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            var appointmentAddedQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'AddedAppointment'";

            var alreadyScheduledQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'AlreadyScheduledAppointment'"; 

            var daysMessageQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = '14DaysMessage'";

            var duplicateEntryQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'DuplicateMessage'";

            var appointmentAddedExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'AddedAppointment'";
            var alreadyScheduledExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'AlreadyScheduledAppointment'";
            var daysMessageExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = '14DaysMessage'";
            var duplicateEntryExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'DuplicateMessage'";

            var appointmentAddedCount = ExecuteScalarQuery(appointmentAddedQuery);
            var alreadyScheduledCount = ExecuteScalarQuery(alreadyScheduledQuery);
            var daysMessageCount = ExecuteScalarQuery(daysMessageQuery);
            var duplicateEntryCount = ExecuteScalarQuery(duplicateEntryQuery);

            var appointmentAddedExceptionCount = ExecuteScalarQuery(appointmentAddedExceptionQuery);
            var alreadyScheduledExceptionCount = ExecuteScalarQuery(alreadyScheduledExceptionQuery);
            var daysMessageExceptionCount = ExecuteScalarQuery(daysMessageExceptionQuery);
            var duplicateEntryExceptionCount = ExecuteScalarQuery(duplicateEntryExceptionQuery);

            var totalHits = appointmentAddedCount + alreadyScheduledCount + daysMessageCount + duplicateEntryCount;
            var totalExceptionsHits = appointmentAddedExceptionCount + alreadyScheduledExceptionCount + daysMessageExceptionCount + duplicateEntryExceptionQuery;

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Add Appointment";
            worksheet.Cells[currentRow, 3].Value = totalHits;
            worksheet.Cells[currentRow, 4].Value = "Appointment Added";
            worksheet.Cells[currentRow, 5].Value = appointmentAddedCount;
            worksheet.Cells[currentRow, 6].Value = totalExceptionsHits;
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Already Scheduled Message";
            worksheet.Cells[currentRow, 5].Value = alreadyScheduledCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "14 Days Message";
            worksheet.Cells[currentRow, 5].Value = daysMessageCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Duplicate Entry";
            worksheet.Cells[currentRow, 5].Value = duplicateEntryCount;
            worksheet.Cells[currentRow, 6].Value = "";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            var addAppointmentCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            addAppointmentCell.Merge = true;
            addAppointmentCell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            addAppointmentCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            addAppointmentCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            addAppointmentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            // Merge "Total Hits" cell across 4 rows
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;

            var exceptionReportedCell = worksheet.Cells[startRow, 6, currentRow - 1, 6];
            exceptionReportedCell.Merge = true;

            var failureCell = worksheet.Cells[startRow, 7, currentRow - 1, 7];
            failureCell.Merge = true;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateReschedule(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            var hoursRescheduleQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'HoursRescheduleAppointment'";

            var rescheduledQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'RescheduleAppointment'";

            var hoursRescheduleExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'HoursRescheduleAppointment'";
            var rescheduledExceptionQuery = @"EXEC WS_PROC_AI_APPOINTMENT_EXCEPTIONS @MethodName = 'RescheduleAppointment'";

            var hoursRescheduleCount = ExecuteScalarQuery(hoursRescheduleQuery);
            var rescheduledCount = ExecuteScalarQuery(rescheduledQuery);

            var hoursRescheduleExceptionCount = ExecuteScalarQuery(hoursRescheduleExceptionQuery);
            var rescheduledExceptionCount = ExecuteScalarQuery(rescheduledExceptionQuery);

            var totalHits = hoursRescheduleCount + rescheduledCount;
            var totalExceptions = 
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Reschedule";
            worksheet.Cells[currentRow, 3].Value = totalHits;
            worksheet.Cells[currentRow, 4].Value = "24 Hours Reschedule";
            worksheet.Cells[currentRow, 5].Value = hoursRescheduleCount;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            var cell = worksheet.Cells[currentRow, 10];
            cell.RichText.Add("Restrict the Patient to ");
            cell.RichText.Add("Reschedule").Bold = true;
            cell.RichText.Add(" due to 24 Hours Check").Bold = false; 
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "Rescheduled";
            worksheet.Cells[currentRow, 5].Value = hoursRescheduleCount;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = 0;
            currentRow++;

            var rescheduleCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            rescheduleCell.Merge = true;
            rescheduleCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            rescheduleCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            // Merge "Total Hits" cell across 3 rows
            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;

            var exceptionReportedCell = worksheet.Cells[startRow, 6, currentRow - 1, 6];
            exceptionReportedCell.Merge = true;

            var failureCell = worksheet.Cells[startRow, 7, currentRow - 1, 7];
            failureCell.Merge = true;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateCancelledAppointments(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            int startRow = currentRow;

            var hoursCancelledAppointmentQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'Cancelled24HoursAppointment'";

            var CancelledAppointmentQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'CancelledAppointment'";

            var hoursCancelledCount = ExecuteScalarQuery(hoursCancelledAppointmentQuery);
            var cancelledCount = ExecuteScalarQuery(CancelledAppointmentQuery);

            var totalHits = hoursCancelledCount + cancelledCount;

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Cancelled";
            worksheet.Cells[currentRow, 3].Value = totalHits;
            worksheet.Cells[currentRow, 4].Value = "Cancelled";
            worksheet.Cells[currentRow, 5].Value = cancelledCount;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;
            worksheet.Cells[currentRow, 8].Value = "";
            currentRow++;

            worksheet.Cells[currentRow, 1].Value = "";
            worksheet.Cells[currentRow, 2].Value = "";
            worksheet.Cells[currentRow, 3].Value = "";
            worksheet.Cells[currentRow, 4].Value = "24 Hours Cancelled";
            worksheet.Cells[currentRow, 5].Value = cancelledCount;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "";
            var cell = worksheet.Cells[currentRow, 10];
            cell.RichText.Add("Restrict the Patient to ");
            cell.RichText.Add("Cancel").Bold = true;
            cell.RichText.Add(" due to 24 Hours Check").Bold = false;
            currentRow++;

            var cancelledCell = worksheet.Cells[startRow, 2, currentRow - 1, 2];
            cancelledCell.Merge = true;
            cancelledCell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cancelledCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

            var totalHitsCell = worksheet.Cells[startRow, 3, currentRow - 1, 3];
            totalHitsCell.Merge = true;

            serialNumber++;

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateLabResults(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            var labResultsQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'LabResult'";
            
            var totalLabResults = ExecuteScalarQuery(labResultsQuery);
            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Lab Results";
            worksheet.Cells[currentRow, 3].Value = totalLabResults;
            worksheet.Cells[currentRow, 4].Value = totalLabResults;
            worksheet.Cells[currentRow, 4, currentRow, 5].Merge = true;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = "No";
            worksheet.Cells[currentRow, 8].Value = 0;
            worksheet.Cells[currentRow, 9].Value = "";

            worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            currentRow++;

            serialNumber++;
            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulateTaskCreation(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            var taskCreationQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = 'TaskCreation'";

            var taskCreationCount = ExecuteScalarQuery(taskCreationQuery);
            
                worksheet.Cells[currentRow, 1].Value = serialNumber;
                worksheet.Cells[currentRow, 2].Value = "Task Creation";
                worksheet.Cells[currentRow, 3].Value = taskCreationCount;
                worksheet.Cells[currentRow, 4].Value = taskCreationCount;
                worksheet.Cells[currentRow, 4, currentRow, 5].Merge = true;
                worksheet.Cells[currentRow, 6].Value = "No";
                worksheet.Cells[currentRow, 7].Value = 0;

                worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 241, 203));

                currentRow++;
                serialNumber++;
            

            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        private PopulateResult PopulatePrescription(ExcelWorksheet worksheet, int currentRow, int serialNumber)
        {
            var prescriptionQuery = @"EXEC sp_Appointment_Tracking_log @MethodName = ''";

            var prescriptionCount = ExecuteScalarQuery(prescriptionQuery);

            worksheet.Cells[currentRow, 1].Value = serialNumber;
            worksheet.Cells[currentRow, 2].Value = "Prescription";
            worksheet.Cells[currentRow, 3].Value = prescriptionCount;
            worksheet.Cells[currentRow, 4].Value = prescriptionCount;
            worksheet.Cells[currentRow, 4, currentRow, 5].Merge = true;
            worksheet.Cells[currentRow, 6].Value = "No";
            worksheet.Cells[currentRow, 7].Value = 0;

            worksheet.Cells[currentRow, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(194, 226, 250));

            currentRow++;

            serialNumber++;


            return new PopulateResult { CurrentRow = currentRow, SerialNumber = serialNumber };
        }

        public async Task<string> SendExcelByEmail()
        {
            try
            {
                GenerateExcelReport();

                string body = $@"
                                <div>
                                    Dear Sir,<br><br>
                                    Please find the attached FDA Agent report of {DateTime.Now.AddDays(-1).ToString("dd-MM-yyyy")}<br><br>
                                    &lt;&lt; FDA Agent API Hits Report - {DateTime.Now.AddDays(-1).ToString("dd-MM-yyyy")} &gt;&gt;<br><br>
                                    <p style='margin-top:1px;'>Note: This is an auto generated email. Please do not reply to this email.</p>
                                </div>";
                //string body = $"<div>Dear Sir,<br><br>Please find the attached FDA Agent report of {DateTime.Now.AddDays(-1).ToString(\"MM-dd-yyyy\")}<br><br>\r\n&lt;&lt; FDA Agent API Hits Report - @{DateTime.Now.AddDays(-1).ToString(\"dd-MM-yyyy\")} &gt;&gt;<br><br>\r\n<p style='margin-top:1px;'>Note: This is an auto generated email. Please do not reply to this email.</p>\r\n</div>";
                string userEmail = ConfigurationManager.AppSettings["MailTo"];
                List<string> listUser = userEmail.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                      .Select(email => email.Trim())
                                                      .ToList();
                if (string.IsNullOrEmpty(userEmail))
                {
                    //return new ResponseResults<string>
                    //{
                    //    Status = false,
                    //    Message = "No email found",
                    //    Data = ""
                    //};
                    return "No email found";
                }

                var bccEmailsString = ConfigurationManager.AppSettings["MailBCC"];
                List<string> listBCC = bccEmailsString.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                                      .Select(email => email.Trim())
                                                      .ToList();

                Email objEmail = new Email
                {
                    body = body,
                    messageTo = string.Join(", ", listUser),
                    subject = "FDA Agent API Calls Report"
                };

                await _emailService.SendEmail(listUser, objEmail, listBCC, _excelFilePath);

                //return new ResponseResults<string>
                //{
                //    Status = true,
                //    Message = "Email sent successfully.",
                //    Data = ""
                //};
                return "Email sent successfully.";
            }
            catch (Exception ex)
            {
                //return new ResponseResults<string>
                //{
                //    Status = false,
                //    Message = $"Something went wrong: {ex.Message}",
                //    Data = ""
                //};
                return $"Something went wrong: {ex.Message}";
            }
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
