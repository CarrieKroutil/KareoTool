using ServiceConnector.KareoApi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ServiceConnector
{
    public class ServiceClient
    {
        private static string OUT_FOLDER_NAME = @"\Output";
        private static string APPOINTMENT_FOLDER_NAME = @"\Appointments";
        private static string CHARGES_FOLDER_NAME = @"\Charges";
        private static string PAYMENTS_FOLDER_NAME = @"\Payments";
        private static string TRANSACTIONS_FOLDER_NAME = @"\Transactions";

        private KareoApi.KareoServices _kareoServices = new KareoServicesClient();
        private RequestHeader _requestHeader;

        public ServiceClient(string customerKey, string apiUser, string apiPassword, string clientVersion)
        {
            _requestHeader = new RequestHeader() { CustomerKey = customerKey, User = apiUser, Password = apiPassword, ClientVersion = clientVersion };
        }

        #region Get & Export Kareo API Data

        #region Get & Export Providers
        public string GetProvidersFromApi()
        {
            ProviderFilter providerFilter = new ProviderFilter() { };
            ProviderFieldsToReturn providerFieldsToReturn = new ProviderFieldsToReturn();
            GetProvidersResp response = null;

            GetProvidersReq request = new GetProvidersReq()
            {
                RequestHeader = _requestHeader,
                Filter = providerFilter,
                Fields = providerFieldsToReturn
            };

            response = _kareoServices.GetProviders(request);
            if (response.ErrorResponse.IsError)
            {
                return response.ErrorResponse.ErrorMessage;
            }
            if (response.Providers == null || response.Providers.Length == 0)
            {
                return "No results.  Check Customer Key is valid in .config file.";
            }

            List<ProviderData> responseData = response.Providers.ToList();

            // Only export active providers
            var data = responseData.Where(p => p.Active == "True");

            ExportProviders(data);
            
            return string.Empty;
        }

        private void ExportProviders(IEnumerable<ProviderData> responseData)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                excelApp.Visible = true;

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
                excelWorksheet.Name = "ProviderList";

                // Establish column headings in cells A1 and B1.
                excelWorksheet.Cells[1, "A"] = "ID";
                excelWorksheet.Cells[1, "B"] = "FirstName";
                excelWorksheet.Cells[1, "C"] = "LastName";
                excelWorksheet.Cells[1, "D"] = "FullName";
                excelWorksheet.Cells[1, "E"] = "Type";
                excelWorksheet.Cells[1, "F"] = "SpecialtyName";
                excelWorksheet.Cells[1, "G"] = "BillingType";
                excelWorksheet.Cells[1, "H"] = "NationalProviderIdentifier";
                excelWorksheet.Cells[1, "I"] = "EmailAddress";
                excelWorksheet.Cells[1, "J"] = "Degree";
                excelWorksheet.Cells[1, "K"] = "CreatedDate";
                excelWorksheet.Cells[1, "L"] = "LastModifiedDate";
                excelWorksheet.Cells[1, "M"] = "Notes";


                var row = 1;
                foreach (var data in responseData)
                {
                    row++;
                    excelWorksheet.Cells[row, "A"] = data.ID;
                    excelWorksheet.Cells[row, "B"] = data.FirstName;
                    excelWorksheet.Cells[row, "C"] = data.LastName;
                    excelWorksheet.Cells[row, "D"] = data.FullName;
                    excelWorksheet.Cells[row, "E"] = data.Type;
                    excelWorksheet.Cells[row, "F"] = data.SpecialtyName;
                    excelWorksheet.Cells[row, "G"] = data.BillingType;
                    excelWorksheet.Cells[row, "H"] = data.NationalProviderIdentifier;
                    excelWorksheet.Cells[row, "I"] = data.EmailAddress;
                    excelWorksheet.Cells[row, "J"] = data.Degree;
                    excelWorksheet.Cells[row, "K"] = data.CreatedDate;
                    excelWorksheet.Cells[row, "L"] = data.LastModifiedDate;
                    excelWorksheet.Cells[row, "M"] = data.Notes;
                }

                string fileName = @"\Providers.xls";
                string currentDirectory = Environment.CurrentDirectory;
                if (!Directory.Exists(currentDirectory + OUT_FOLDER_NAME))
                {
                    Directory.CreateDirectory(currentDirectory + OUT_FOLDER_NAME);
                }

                excelApp.DisplayAlerts = false;
                excelApp.ActiveWorkbook.SaveAs(currentDirectory + OUT_FOLDER_NAME + fileName, Excel.XlFileFormat.xlWorkbookNormal);

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion

        #region Get & Export Patients
        public string GetPatientsFromApi()
        {
            PatientFilter patientFilter = new PatientFilter() { };
            PatientFieldsToReturn patientFieldsToReturn = new PatientFieldsToReturn();
            GetPatientsResp response = null;

            GetPatientsReq request = new GetPatientsReq()
            {
                RequestHeader = _requestHeader,
                Filter = patientFilter,
                Fields = patientFieldsToReturn
            };

            response = _kareoServices.GetPatients(request);
            if (response.ErrorResponse.IsError)
            {
                return response.ErrorResponse.ErrorMessage;
            }
            if (response.Patients == null || response.Patients.Length == 0)
            {
                return "No results.  Check Customer Key is valid in .config file.";
            }

            List<PatientData> responseData = response.Patients.ToList();

            // Only export active patients
            //var data = responseData.Where(p => p.Active == "True");

            ExportPatients(responseData);

            return string.Empty;
        }

        private void ExportPatients(List<PatientData> responseData)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                excelApp.Visible = true;

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
                excelWorksheet.Name = "PatientList";

                // Establish column headings in cells A1 and B1.
                excelWorksheet.Cells[1, "A"] = "ID";
                excelWorksheet.Cells[1, "B"] = "FirstName"; // Do we really need PII ?
                excelWorksheet.Cells[1, "C"] = "LastName";  // Do we really need PII ?
                excelWorksheet.Cells[1, "D"] = "FullName";  // Do we really need PII ?
                excelWorksheet.Cells[1, "E"] = "DefaultRenderingProviderId";
                excelWorksheet.Cells[1, "F"] = "DefaultRenderingProviderFullName";
                excelWorksheet.Cells[1, "G"] = "DefaultServiceLocationBillingName";
                excelWorksheet.Cells[1, "H"] = "CreatedDate";
                excelWorksheet.Cells[1, "I"] = "LastEncounterDate";
                excelWorksheet.Cells[1, "J"] = "LastAppointmentDate";
                excelWorksheet.Cells[1, "K"] = "InsuranceBalance";
                excelWorksheet.Cells[1, "L"] = "InsurancePayments";
                excelWorksheet.Cells[1, "M"] = "LastPaymentDate";
                excelWorksheet.Cells[1, "N"] = "PatientBalance";
                excelWorksheet.Cells[1, "O"] = "PatientPayments";
                excelWorksheet.Cells[1, "P"] = "Adjustments";
                excelWorksheet.Cells[1, "Q"] = "Charges";
                excelWorksheet.Cells[1, "R"] = "PrimaryInsurancePolicyPlanID";
                excelWorksheet.Cells[1, "S"] = "PrimaryInsurancePolicyPlanName";
                excelWorksheet.Cells[1, "T"] = "PrimaryInsurancePolicyCopay";
                excelWorksheet.Cells[1, "U"] = "PrimaryInsurancePolicyDeductible";
                excelWorksheet.Cells[1, "V"] = "SecondaryInsurancePolicyPlanID";
                excelWorksheet.Cells[1, "W"] = "SecondaryInsurancePolicyCompanyName";
                excelWorksheet.Cells[1, "X"] = "SecondaryInsurancePolicyCopay";
                excelWorksheet.Cells[1, "Y"] = "SecondaryInsurancePolicyDeductible";

                var row = 1;
                foreach (var data in responseData)
                {
                    row++;
                    excelWorksheet.Cells[row, "A"] = data.ID;
                    excelWorksheet.Cells[row, "B"] = data.FirstName;
                    excelWorksheet.Cells[row, "C"] = data.LastName;
                    excelWorksheet.Cells[row, "D"] = data.PatientFullName;
                    excelWorksheet.Cells[row, "E"] = data.DefaultRenderingProviderId;
                    excelWorksheet.Cells[row, "F"] = data.DefaultRenderingProviderFullName;
                    excelWorksheet.Cells[row, "G"] = data.DefaultServiceLocationBillingName;
                    excelWorksheet.Cells[row, "H"] = data.CreatedDate;
                    excelWorksheet.Cells[row, "I"] = data.LastEncounterDate;
                    excelWorksheet.Cells[row, "J"] = data.LastAppointmentDate;
                    excelWorksheet.Cells[row, "K"] = data.InsuranceBalance;
                    excelWorksheet.Cells[row, "L"] = data.InsurancePayments;
                    excelWorksheet.Cells[row, "M"] = data.LastPaymentDate;
                    excelWorksheet.Cells[row, "N"] = data.PatientBalance;
                    excelWorksheet.Cells[row, "O"] = data.PatientPayments;
                    excelWorksheet.Cells[row, "P"] = data.Adjustments;
                    excelWorksheet.Cells[row, "Q"] = data.Charges;
                    excelWorksheet.Cells[row, "R"] = data.PrimaryInsurancePolicyPlanID;
                    excelWorksheet.Cells[row, "S"] = data.PrimaryInsurancePolicyPlanName;
                    excelWorksheet.Cells[row, "T"] = data.PrimaryInsurancePolicyCopay;
                    excelWorksheet.Cells[row, "U"] = data.PrimaryInsurancePolicyDeductible;
                    excelWorksheet.Cells[row, "V"] = data.SecondaryInsurancePolicyPlanID;
                    excelWorksheet.Cells[row, "W"] = data.SecondaryInsurancePolicyCompanyName;
                    excelWorksheet.Cells[row, "X"] = data.SecondaryInsurancePolicyCopay;
                    excelWorksheet.Cells[row, "Y"] = data.SecondaryInsurancePolicyDeductible;
                }


                string fileName = @"\Patients.xls";
                string currentDirectory = Environment.CurrentDirectory;
                if (!Directory.Exists(currentDirectory + OUT_FOLDER_NAME))
                {
                    Directory.CreateDirectory(currentDirectory + OUT_FOLDER_NAME);
                }

                excelApp.DisplayAlerts = false;
                excelApp.ActiveWorkbook.SaveAs(currentDirectory + OUT_FOLDER_NAME + fileName, Excel.XlFileFormat.xlWorkbookNormal);

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion Get & Export Patients

        #region Get & Export Transations
        public string GetTransactionsFromApi(string fromDate, string toDate)
        {
            TransactionFilter transactionFilter = new TransactionFilter() { FromServiceDate = fromDate, ToServiceDate = toDate };
            TransactionFieldsToReturn transactionFieldsToReturn = new TransactionFieldsToReturn();
            GetTransactionsResp response = null;

            GetTransactionsReq request = new GetTransactionsReq()
            {
                RequestHeader = _requestHeader,
                Filter = transactionFilter,
                Fields = transactionFieldsToReturn
            };

            response = _kareoServices.GetTransactions(request);
            if (response.ErrorResponse.IsError)
            {
                return response.ErrorResponse.ErrorMessage;
            }
            if (response.Transactions == null || response.Transactions.Length == 0)
            {
                return "No results.  Check Customer Key is valid in .config file.";
            }

            List<TransactionData> responseData = response.Transactions.ToList();

            // Only export active transactions
            //var data = responseData.Where(p => p.Active == "True");

            string fileName = @"\Transactions_" + fromDate + "_" + toDate + ".xls";
            ExportTransactions(responseData, fileName);

            return string.Empty;
        }

        private void ExportTransactions(List<TransactionData> responseData, string exportToFileName)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                excelApp.Visible = true;

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
                excelWorksheet.Name = "TransactionList";

                // Establish column headings in cells A1 and B1.
                excelWorksheet.Cells[1, "A"] = "ID";
                excelWorksheet.Cells[1, "B"] = "PatientID";
                excelWorksheet.Cells[1, "C"] = "PatientFullName";
                excelWorksheet.Cells[1, "D"] = "TransactionDate";
                excelWorksheet.Cells[1, "E"] = "PostingDate";
                excelWorksheet.Cells[1, "F"] = "ServiceDate";
                excelWorksheet.Cells[1, "G"] = "Type";
                excelWorksheet.Cells[1, "H"] = "ProcedureCode";
                excelWorksheet.Cells[1, "I"] = "Amount";
                excelWorksheet.Cells[1, "J"] = "Description";
                excelWorksheet.Cells[1, "K"] = "InsuranceOrder";
                excelWorksheet.Cells[1, "L"] = "InsuranceID";
                excelWorksheet.Cells[1, "M"] = "InsuranceCompanyName";
                excelWorksheet.Cells[1, "N"] = "InsurancePlanName";
                excelWorksheet.Cells[1, "O"] = "LastModifiedDate";

                var row = 1;
                foreach (var data in responseData)
                {
                    row++;
                    excelWorksheet.Cells[row, "A"] = data.ID;
                    excelWorksheet.Cells[row, "b"] = data.PatientID;
                    excelWorksheet.Cells[row, "c"] = data.PatientFullName;
                    excelWorksheet.Cells[row, "d"] = data.TransactionDate;
                    excelWorksheet.Cells[row, "e"] = data.PostingDate;
                    excelWorksheet.Cells[row, "f"] = data.ServiceDate;
                    excelWorksheet.Cells[row, "g"] = data.Type;
                    excelWorksheet.Cells[row, "h"] = data.ProcedureCode;
                    excelWorksheet.Cells[row, "i"] = data.Amount;
                    excelWorksheet.Cells[row, "j"] = data.Description;
                    excelWorksheet.Cells[row, "k"] = data.InsuranceOrder;
                    excelWorksheet.Cells[row, "l"] = data.InsuranceID;
                    excelWorksheet.Cells[row, "m"] = data.InsuranceCompanyName;
                    excelWorksheet.Cells[row, "n"] = data.InsurancePlanName;
                    excelWorksheet.Cells[row, "o"] = data.LastModifiedDate;
                }

                string currentDirectory = Environment.CurrentDirectory;
                if (!Directory.Exists(currentDirectory + OUT_FOLDER_NAME))
                {
                    Directory.CreateDirectory(currentDirectory + OUT_FOLDER_NAME);
                }
                if (!Directory.Exists(currentDirectory + OUT_FOLDER_NAME + TRANSACTIONS_FOLDER_NAME))
                {
                    Directory.CreateDirectory(currentDirectory + OUT_FOLDER_NAME + TRANSACTIONS_FOLDER_NAME);
                }

                excelApp.DisplayAlerts = false;
                excelApp.ActiveWorkbook.SaveAs(currentDirectory + OUT_FOLDER_NAME + TRANSACTIONS_FOLDER_NAME + exportToFileName, Excel.XlFileFormat.xlWorkbookNormal);

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion Get & Export Transactions

        #endregion Get & Export Kareo API Data
    }
}
