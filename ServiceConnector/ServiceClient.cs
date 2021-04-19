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
            var data = responseData.Where(p => p.Active == "True").ToList();

            return ExportProviders(data);
        }

        private string ExportProviders(List<ProviderData> responseData)
        {
            var exportSettings = new ExcelExporter.ExportSettings()
            {
                ExportDirectoryName = OUT_FOLDER_NAME,
                EnableExportToSubFolder = false,
                ExportToSubFolderName = String.Empty,
                ExportToFileName = @"\Providers.xls"
            };

            return new ExcelExporter().ExportToExcel(responseData, exportSettings);

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

            return ExportPatients(responseData);
        }

        private string ExportPatients(List<PatientData> responseData)
        {
            var exportSettings = new ExcelExporter.ExportSettings()
            {
                ExportDirectoryName = OUT_FOLDER_NAME,
                EnableExportToSubFolder = false,
                ExportToSubFolderName = String.Empty,
                ExportToFileName = @"\Patients.xls"
            };

            return new ExcelExporter().ExportToExcel(responseData, exportSettings);
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

            string fileName = @"\Transactions_" + fromDate + "_" + toDate + ".xls";
            return ExportTransactions(responseData, fileName);
        }

        private string ExportTransactions(List<TransactionData> responseData, string exportToFileName)
        {
            var exportSettings = new ExcelExporter.ExportSettings()
            {
                ExportDirectoryName = OUT_FOLDER_NAME,
                EnableExportToSubFolder = true,
                ExportToSubFolderName = TRANSACTIONS_FOLDER_NAME,
                ExportToFileName = exportToFileName
            };

            return new ExcelExporter().ExportToExcel(responseData, exportSettings);
        }
        #endregion Get & Export Transactions

        #endregion Get & Export Kareo API Data
    }
}
