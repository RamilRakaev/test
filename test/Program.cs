using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;

namespace test
{
    public class Program
    {
        const string ApiKey = "AIzaSyD2C_-YNu5UONpXBXLI3en9XtlGOvO_gtI";
        const string CalendarId = "ru.russian#holiday@group.v.calendar.google.com";
        const string path = "C:\\Users\\Admin\\source\\repos\\test\\test\\credentials.json";

        static async Task Main(string[] args)
        {
            //var service = new CalendarService(new BaseClientService.Initializer()
            //{
            //    ApiKey = ApiKey,
            //    ApplicationName = "App"
            //});

            //var request = service.Events.List(CalendarId);
            //request.Fields = "items(summary,start,end)";
            //var response = await request.ExecuteAsync();

            //foreach (var item in response.Items)
            //{
            //    Console.WriteLine($"Holiday: {item.Summary} start: {item.Start} end: {item.End}");
            //}
            var google = new GoogleCalendar(path);
            //google.ShowUpCommingEvent();
            //google.CreateEvent();
            Console.WriteLine();
            google.ShowUpCommingEvent();
            Console.ReadLine();
        }
    }

    class GoogleCalendar
    {
        public static string[] Scopes = { CalendarService.Scope.Calendar };
        public static string ApplicationName = "CalendarConsole";
        const string ApiKey = "AIzaSyD2C_-YNu5UONpXBXLI3en9XtlGOvO_gtI";
        const string CalendarId = "ru.russian#holiday@group.v.calendar.google.com";

        private string CredentialsPath = string.Empty;

        public GoogleCalendar(string credentialsPath)
        {
            CredentialsPath = credentialsPath;
        }

        public void ShowUpCommingEvent()
        {
            UserCredential credential = GetCredential(UserRole.User);

            // Creat Google Calendar API service.
            CalendarService service = GetService(credential);

            // Define parameters of request
            EventsResource.ListRequest request = service.Events.List(CalendarId);
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events.
            Events events = request.Execute();

            // Print upcomming events
            Console.WriteLine("Upcomming events:");
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    Console.WriteLine("{0} ({1})", eventItem.Summary, when);
                }
            }
            else
            {
                Console.WriteLine("Nothing.");
            }
        }

        public void CreateEvent()
        {
            UserCredential credential = GetCredential(UserRole.Admin);
            CalendarService service = GetService(credential);

            Event newEvent = new Event()
            {
                Summary = "LIFO Event",
                Start = new EventDateTime() { DateTime = new DateTime(2018, 10, 20) },
                End = new EventDateTime() { DateTime = new DateTime(2018, 10, 21) },
            };

            newEvent = service.Events.Insert(newEvent, CalendarId).Execute();
            Console.WriteLine($"{newEvent.HtmlLink}");
        }

        private CalendarService GetService(UserCredential credential)
        {
            // Creat Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        private UserCredential GetCredential(UserRole userRole)
        {
            UserCredential credential;
            using (var stream =
                new FileStream(CredentialsPath, FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                Scopes,
                userRole.ToString(),
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;

                Console.WriteLine($"Credential file saved to: {credPath}");
            }

            return credential;
        }
    }
    public enum UserRole
    {
        Admin,
        User
    }

}

