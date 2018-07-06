using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleCalendarSample
{

    class Program
    {
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");

                Console.WriteLine(credPath);
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "admin",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            Event myEvent = new Event
            {
                Summary = "Appointment",
                Location = "Somewhere",
                Start = new EventDateTime()
                {
                    DateTime = new DateTime(2018, 7, 7, 10, 0, 0),
                    TimeZone = "America/Los_Angeles"
                },
                End = new EventDateTime()
                {
                    DateTime = new DateTime(2018, 7, 7, 10, 30, 0),
                    TimeZone = "America/Los_Angeles"
                },
                Recurrence = new String[] {
                                     "RRULE:FREQ=WEEKLY;BYDAY=MO"
                 },
                Attendees = new List<EventAttendee>(){
                                   new EventAttendee() { Email = "mlehal@advantageds.com" }
                              }
            };

            Event recurringEvent = service.Events.Insert(myEvent, "primary").Execute();


            Console.Read();
        }
    }
}
