using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace PraticeManagementExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            var configHelper = new ConfigHelper();
            var apiCreds = configHelper.GetApiCreds();
            var settings = configHelper.GetEnabledSettings();

            var client = new ServiceConnector.ServiceClient(apiCreds.customerKeyConfig, apiCreds.apiUserConfig, apiCreds.apiPasswordConfig, apiCreds.clientVersionConfig);

            if (settings.AreProvidersEnabled)
            {
                Console.WriteLine("Calling Api for providers and exporting data to excel...");
                client.GetProvidersFromApi();
                Console.WriteLine("Providers export completed successful.");
            }

            if (settings.ArePatientsEnabled)
            {
                Console.WriteLine("Calling Api for patients and exporting data to excel...");
                client.GetPatientsFromApi();
                Console.WriteLine("Patients export completed successful.");
            }

            Console.WriteLine("Done");
            Console.Read();
        }
    }
}
