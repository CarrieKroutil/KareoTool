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

        /// <summary>
        /// Settings for which data endpoints to export.
        /// </summary>
        public EnabledSettings()
        {
            // Providers
            string enableProvidersConfig = ConfigurationManager.AppSettings["EnableProviders"];
            if (!string.IsNullOrWhiteSpace(enableProvidersConfig))
            {
                try
                {
                    AreProvidersEnabled = bool.Parse(enableProvidersConfig);
                    Console.WriteLine($"Config: EnableProviders = '{AreProvidersEnabled}'");
                }
                catch (Exception)
                {

                    throw;
                }
            }
            else
            {
                Console.WriteLine("Missing Config: EnableProviders");
            }

            // Patients
            string enablePatientsConfig = ConfigurationManager.AppSettings["EnablePatients"];
            if (!string.IsNullOrWhiteSpace(enablePatientsConfig))
            {
                try
                {
                    ArePatientsEnabled = bool.Parse(enablePatientsConfig);
                    Console.WriteLine($"Config: EnablePatients = '{ArePatientsEnabled}'");
                }
                catch (Exception)
                {

                    throw;
                }
            }
            else
            {
                Console.WriteLine("Missing Config: EnablePatients");
            }
        }
    }
}
