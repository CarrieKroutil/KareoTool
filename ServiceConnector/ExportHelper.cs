using Microsoft.Office.Interop.Excel;
using ServiceConnector.KareoApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace ServiceConnector
{
    internal class ExportHelper
    {
        internal string SetWorksheet<T>(List<T> responseData, Excel.Worksheet excelWorksheet)
        {
            switch (typeof(T).Name)
            {
                case nameof(ProviderData):
                    ExportProviderData(responseData, excelWorksheet);
                    break;
                case nameof(PatientData):
                    ExportPatientData(responseData, excelWorksheet);
                    break;
                case nameof(AppointmentData):
                    ExportAppointmentData(responseData, excelWorksheet);
                    break;
                case nameof(ChargeData):
                    ExportChargeData(responseData, excelWorksheet);
                    break;
                case nameof(PaymentData):
                    ExportPaymentData(responseData, excelWorksheet);
                    break;
                case nameof(TransactionData):
                    ExportTransactionData(responseData, excelWorksheet);
                    break;
                case nameof(EncounterDetailsData):
                    ExportEncounterDetailsData(responseData, excelWorksheet);
                    break;
                default:
                    return $"Type - {typeof(T).Name} not impleted to export";
            }
            return string.Empty;
        }

        #region Export Specific Type Implemenation

        private void ExportProviderData<T>(List<T> responseData, Worksheet excelWorksheet)
        {
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
            foreach (var data in responseData as List<ProviderData>)
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
        }

        private void ExportPatientData<T>(List<T> responseData, Worksheet excelWorksheet)
        {
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
            foreach (var data in responseData as List<PatientData>)
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
        }

        private void ExportAppointmentData<T>(List<T> responseData, Worksheet excelWorksheet)
        {
            throw new NotImplementedException();
        }

        private void ExportChargeData<T>(List<T> responseData, Worksheet excelWorksheet)
        {
            throw new NotImplementedException();
        }

        private void ExportPaymentData<T>(List<T> responseData, Worksheet excelWorksheet)
        {
            throw new NotImplementedException();
        }

        private void ExportEncounterDetailsData<T>(List<T> responseData, Worksheet excelWorksheet)
        {
            throw new NotImplementedException();
        }

        private void ExportTransactionData<T>(List<T> responseData, Excel.Worksheet excelWorksheet)
        {            
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
            foreach (var data in responseData as List<TransactionData>)
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
        }

        #endregion
    }
}
