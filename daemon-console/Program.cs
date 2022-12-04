using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace ms_graph_app_auth
{

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var date  = DateTime.Now;//12/1/2022 3:51:34 PM
                RunAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }


        private static async Task RunAsync()
        {
            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");
            // Even if this is a console application here, a daemon application is a confidential client application
            IConfidentialClientApplication app;
            // Even if this is a console application here, a daemon application is a confidential client application
            app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                .WithClientSecret(config.ClientSecret)
                .WithAuthority(new Uri(config.Authority))
                .Build();

            app.AddInMemoryTokenCache();

            // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
            // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
            // a tenant administrator. 
            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; // Generates a scope -> "https://graph.microsoft.com/.default"

            // Call MS graph using the Graph SDK
            await CallMSGraphUsingGraphSDK(app, scopes, config.email);

        }


        /// <summary>
        /// The following example shows how to initialize the MS Graph SDK
        /// </summary>
        /// <param name="app"></param>
        /// <param name="scopes"></param>
        /// <returns></returns>
        private static async Task CallMSGraphUsingGraphSDK(IConfidentialClientApplication app, string[] scopes, string email)
        {
            // Prepare an authenticated MS Graph SDK client
            GraphServiceClient graphServiceClient = GetAuthenticatedGraphClient(app, scopes);

            //AuthenticationResult authenticationResult = await app.AcquireTokenForClient(scopes).ExecuteAsync();

            List<User> allUsers = new List<User>();

            try
            {

                //IGraphServiceUsersCollectionPage users = await graphServiceClient.Users.Request().GetAsync();
                //Console.WriteLine($"Found {users.Count()} users in the tenant");


                NewEvent newEvent = NewEvent.ReadFromJsonFile("eventbody.json");

                var Period = newEvent.Period;
                var startDate = newEvent.StartDate.AddHours(newEvent.StartTime.Hour).AddMinutes(newEvent.StartTime.Minute);
                var endDate = newEvent.StartDate.AddHours(newEvent.EndTime.Hour).AddMinutes(newEvent.EndTime.Minute);


                // Create a Graph event with the required fields
                var graphEvent = new Event
                {
                    Subject = newEvent.Subject,
                    Start = new DateTimeTimeZone
                    {
                        DateTime = startDate.ToString("o"),
                        // Use the user's time zone
                        TimeZone = "Arab Standard Time"
                    },
                    End = new DateTimeTimeZone
                    {
                        DateTime = endDate.ToString("o"),
                        // Use the user's time zone
                        TimeZone = "Arab Standard Time"
                    },
                    Location = new Location
                    {
                        DisplayName = newEvent.MeetingLocation
                    }
                };

                if (IsOnline(newEvent.Classification))
                {
                    graphEvent.IsOnlineMeeting = true;

                    graphEvent.OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness;
                }

                if (newEvent.EventRecurring)
                {
                    graphEvent.Recurrence = GetRecurrenceObject(Period, newEvent.StartDate, newEvent.EndDate);
                }
                // Add body if present
                if (!string.IsNullOrEmpty(newEvent.Body))
                {
                    graphEvent.Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = newEvent.Body
                    };
                }

                // Add attendees if present
                if (newEvent.Attendees != null)
                {

                    var attendeeList = new List<Attendee>();
                    foreach (var attendeeEmail in newEvent.Attendees)
                    {
                        attendeeList.Add(new Attendee
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = attendeeEmail
                            },
                            Type = AttendeeType.Required
                        });
                    }

                    graphEvent.Attendees = attendeeList;

                }



                //Event response = await graphServiceClient.Me.Events
                //.Request()
                //.AddAsync(graphEvent);

                //for our configrations mamen@1jqkmk.onmicrosoft.com
                //for nwc configrations use fmossa.c@nwc.com.sa or
                Event response2 = await graphServiceClient.Users[email].Events
                    .Request()
                    .AddAsync(graphEvent);

            }
            catch (ServiceException e)
            {
                Console.WriteLine("Can not create event: " + $"{e}");
            }

        }

  
        /// <summary>
        /// An example of how to authenticate the Microsoft Graph SDK using the MSAL library
        /// </summary>
        /// <returns></returns>
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
        {            

            GraphServiceClient graphServiceClient =
                    new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                        AuthenticationResult result = await app.AcquireTokenForClient(scopes)
                            .ExecuteAsync();

                        // Add the access token in the Authorization header of the API request.
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));

            return graphServiceClient;
        }


        private static PatternedRecurrence GetRecurrenceObject(string Period, DateTime startDate, DateTime endDate)
        {
            PatternedRecurrence PatternedRecurrence = new PatternedRecurrence();

            if (Period == "Every Day" || Period == "كل يوم")
            {
                PatternedRecurrence = new PatternedRecurrence()
                {
                    Pattern = new RecurrencePattern()
                    {
                        Type = RecurrencePatternType.Daily,
                        Interval = 1
                    },
                    Range = new RecurrenceRange()
                    {
                        Type = RecurrenceRangeType.EndDate,
                        StartDate = new Date(startDate.Year, startDate.Month, startDate.Day),
                        EndDate = new Date(endDate.Year, endDate.Month, endDate.Day)
                    }
                };

            }
            else if (Period == "Every Week" || Period == "كل اسبوع")
            {
                var DaysOfWeek = new List<Microsoft.Graph.DayOfWeek>();

                DaysOfWeek.Add(getGraphDayOfWeek(startDate.DayOfWeek.ToString()));

                //while (endDate >= startDate)
                //{
                //    DaysOfWeek.Add(getGraphDayOfWeek(startDate.DayOfWeek.ToString()));
                //    startDate = startDate.AddDays(1);
                //}
                //type, interval, daysOfWeek, firstDayOfWeek
                PatternedRecurrence = new PatternedRecurrence()
                {
                    Pattern = new RecurrencePattern()
                    {
                        Type = RecurrencePatternType.Weekly,
                        Interval = 1,
                        DaysOfWeek = DaysOfWeek.Distinct(),
                        FirstDayOfWeek = Microsoft.Graph.DayOfWeek.Saturday
                    },
                    Range = new RecurrenceRange()
                    {
                        Type = RecurrenceRangeType.EndDate,
                        StartDate = new Date(startDate.Year, startDate.Month, startDate.Day),
                        EndDate = new Date(endDate.Year, endDate.Month, endDate.Day)
                    }
                };
            }
            else if (Period == "Every Month" || Period == "كل شهر")
            {
                //type, interval, dayOfMonth

                PatternedRecurrence = new PatternedRecurrence()
                {
                    Pattern = new RecurrencePattern()
                    {
                        Type = RecurrencePatternType.AbsoluteMonthly,
                        Interval = 1,
                        DayOfMonth = startDate.Day
                    },
                    Range = new RecurrenceRange()
                    {
                        Type = RecurrenceRangeType.NoEnd,
                        StartDate = new Date(startDate.Year, startDate.Month, startDate.Day),
                        EndDate = new Date(endDate.Year, endDate.Month, endDate.Day)
                    }
                };
            }

            return PatternedRecurrence;
        }
        private static Microsoft.Graph.DayOfWeek getGraphDayOfWeek(string dayName)
        {
            switch (dayName)
            {
                case "Saturday":
                    return Microsoft.Graph.DayOfWeek.Saturday;
                    break;
                case "Sunday":
                    return Microsoft.Graph.DayOfWeek.Sunday;
                    break;
                case "Monday":
                    return Microsoft.Graph.DayOfWeek.Monday;
                    break;
                case "Tuesday":
                    return Microsoft.Graph.DayOfWeek.Tuesday;
                    break;
                case "Wednesday":
                    return Microsoft.Graph.DayOfWeek.Wednesday;
                    break;
                case "Thursday":
                    return Microsoft.Graph.DayOfWeek.Thursday;
                    break;
                case "Friday":
                    return Microsoft.Graph.DayOfWeek.Friday;
                    break;
                default:
                    return Microsoft.Graph.DayOfWeek.Friday;
                    break;
            }
        }
        private static bool IsMeeting(string requestType)
        {
            if (requestType == "Meeting" || requestType == "اجتماع")
            {
                return true;

            }

            return false;
        }
        private static bool IsOnline(string requestClassification)
        {
            if (requestClassification == "Online" || requestClassification == "عبر الإنترنت")
            {
                return true;

            }

            return false;

        }

    }
}
