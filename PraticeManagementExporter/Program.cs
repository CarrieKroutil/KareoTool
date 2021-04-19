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
            try
            {
                // Extract Settings from App.Config:
                var configHelper = new ConfigHelper();
                var apiCreds = configHelper.GetApiCreds();
                var settings = configHelper.GetEnabledSettings();

                // Instantiate Kareo API ServiceClient:
                var client = new ServiceConnector.ServiceClient(apiCreds.customerKeyConfig, apiCreds.apiUserConfig, apiCreds.apiPasswordConfig, apiCreds.clientVersionConfig);

                // Call Api for Providers and Extract to Excel:
                if (settings.AreProvidersEnabled)
                {
                    Console.WriteLine("\nCalling Api for providers and exporting data to excel...");
                    string response = client.GetProvidersFromApi();
                    if (!string.IsNullOrWhiteSpace(response))
                    {
                        Console.WriteLine($"\nApi Message: {response}\n");
                    }
                    else
                    {
                        Console.WriteLine("Providers export completed successful.\n");
                    }
                }

                // Call Api for Patients and Extract to Excel:
                if (settings.ArePatientsEnabled)
                {
                    Console.WriteLine("\nCalling Api for patients and exporting data to excel...");
                    string response = client.GetPatientsFromApi();
                    if (!string.IsNullOrWhiteSpace(response))
                    {
                        Console.WriteLine($"\nApi Message: {response}\n");
                    }
                    else
                    {
                        Console.WriteLine("Patients export completed successful.\n");
                    }
                }

                // Call Api for Appointments and Extract to Excel:
                if (settings.AreAppointmentsEnabled)
                {
                    Console.WriteLine("\nAppointments have not been coded yet...");

                }

                // Call Api for Charges and Extract to Excel:
                if (settings.AreChargesEnabled)
                {
                    Console.WriteLine("\nAppointments have not been coded yet...");

                }

                // Call Api for Payments and Extract to Excel:
                if (settings.ArePaymentsEnabled)
                {
                    Console.WriteLine("\nAppointments have not been coded yet...");

                }

                // Call Api for Transactions and Extract to Excel:
                if (settings.AreTransactionsEnabled)
                {
                    Console.WriteLine("\nCalling Api for transactions and exporting data to excel...");
                    string response = client.GetTransactionsFromApi(configHelper.GetTransactionsFromServiceDate, configHelper.GetTransactionsToServiceDate);
                    if (!string.IsNullOrWhiteSpace(response))
                    {
                        Console.WriteLine($"\nApi Message: {response}\n");
                    }
                    else
                    {
                        Console.WriteLine("Transactions export completed successful.\n");
                    }
                }

                // Call Api for Encounters and Extract to Excel:
                /* NOTE: This ties all the APIs together with EncounterID, and contains: PatientID, AppointmentID, RenderingProviderID.
                    Then charges & payments are related to the encounter's "BatchNumber".
                    Transactions do not have any key to tie back to encounters, other than the PatientID and service date.     
                 */ 
                if (settings.AreEncountersEnabled)
                {
                    Console.WriteLine("\nEncounters have not been coded yet...");

                }
            }
            catch (Exception err)
            {
                Console.WriteLine("\n\nIssues with processing occurred... process stopped.");
                Console.WriteLine($"\n\n Error Message:\n {err.Message} \n\n Stack Trace:\n {err.StackTrace}");
                Console.Read();
            }

            // TODO: Add logger implemenation to write every time cw is written to help track reported issues.
            Console.WriteLine("Done");
            Console.Read();
        }
    }
}
