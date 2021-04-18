using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PraticeManagementExporter
{
    public class ConfigHelper
    {
        public ApiConnectionCreds GetApiCreds()
        {
            return new ApiConnectionCreds();
        }

        public EnabledSettings GetEnabledSettings()
        {
            return new EnabledSettings();
        }

        public string GetTransactionsFromServiceDate {
            get
            {
                var value = ConfigurationManager.AppSettings["GetTransactionsFromServiceDate"];
                if (string.IsNullOrWhiteSpace(value))
                {
                    Console.WriteLine("Missing Config: GetTransactionsFromServiceDate. Default value of yesterday will be used.");
                    value = DateTime.Today.AddDays(-1).ToString();
                }
                return value;
            }
        }
        public string GetTransactionsToServiceDate
        {
            get
            {
                var value = ConfigurationManager.AppSettings["GetTransactionsToServiceDate"];
                if (string.IsNullOrWhiteSpace(value))
                {
                    Console.WriteLine("Missing Config: GetTransactionsToServiceDate. Default value of today will be used.");
                    value = DateTime.Today.ToString();
                }
                return value;
            }
        }
    }

    public class ApiConnectionCreds
    {
        public string customerKeyConfig = ConfigurationManager.AppSettings["CustomerKey"];
        public string apiUserConfig = ConfigurationManager.AppSettings["ApiUser"];
        public string apiPasswordConfig = ConfigurationManager.AppSettings["ApiPassword"];
        public string clientVersionConfig = ConfigurationManager.AppSettings["ClientVersion"];
    }

    public class EnabledSettings
    {
        public bool AreProvidersEnabled { get; private set; }
        public bool ArePatientsEnabled { get; private set; }
        public bool AreAppointmentsEnabled { get; private set; }
        public bool AreChargesEnabled { get; private set; }
        public bool ArePaymentsEnabled { get; private set; }
        public bool AreTransactionsEnabled { get; private set; }
        public bool AreEncountersEnabled { get; private set; }

        /// <summary>
        /// Settings for which data endpoints to export.
        /// </summary>
        public EnabledSettings()
        {
            // Providers
            AreProvidersEnabled = GetEnabledSettingValue("EnableProviders");

            // Patients
            ArePatientsEnabled = GetEnabledSettingValue("EnablePatients");

            // Appointments
            AreTransactionsEnabled = GetEnabledSettingValue("EnableAppointments");

            // Charges
            AreTransactionsEnabled = GetEnabledSettingValue("EnableCharges");

            // Payments
            AreTransactionsEnabled = GetEnabledSettingValue("EnablePayments");

            // Transations
            AreTransactionsEnabled = GetEnabledSettingValue("EnableTransactions");

            // Encounters
            AreEncountersEnabled = GetEnabledSettingValue("EnableEncounters");
        }

        private bool GetEnabledSettingValue(string configName)
        {
            string configValue = ConfigurationManager.AppSettings[configName];
            bool enabledSetting = false;
            if (!string.IsNullOrWhiteSpace(configValue))
            {
                try
                {
                    enabledSetting = bool.Parse(configValue);
                    Console.WriteLine($"Config: {configName} = '{enabledSetting}'");
                }
                catch (Exception err)
                {
                    Console.WriteLine($"Config value for {configName}: {configValue} \nError Message: {err.Message}");
                    throw;
                }
            }
            else
            {
                Console.WriteLine($"Missing Config: {configName}. Default value of false will be used.");
            }

            return enabledSetting;
        }
    }
}
