using NLog;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

public static class Functions
{
    private static NLog.Logger l = LogManager.GetCurrentClassLogger();
    public class Feiertag
    {
        public string Subject { get; set; }
        public string StartDatetime { get; set; }
        public string EndDateTime { get; set; }
    }

    public static List<Feiertag> ReadCSV(string file, string today)
    //load events from csv-file to an List of "Feiertag"-Objects.
    {
        try
        {
            string[] lines = System.IO.File.ReadAllLines(file);
            List<Feiertag> FeiertageAusCsv = new List<Feiertag>();
            for (int i = 1; i < lines.Length; i++) //skips first row!
            {
                Feiertag tempFeiertag = new Feiertag();
                string[] values = lines[i].Split(',');
                tempFeiertag.Subject = values[0].Trim(new Char[] { '"' });
                tempFeiertag.StartDatetime = values[1].Trim(new Char[] { '"' });
                tempFeiertag.EndDateTime = values[2].Trim(new Char[] { '"' });
                if (DateTime.Compare(Convert.ToDateTime(tempFeiertag.StartDatetime), Convert.ToDateTime(today)) >= 0) //only add future events to list.
                {
                    FeiertageAusCsv.Add(tempFeiertag);
                }
            }
            return FeiertageAusCsv;
        }
        catch (FileNotFoundException e)
        {
            l.Fatal("!-- EXIT --! at ReadCSV with FileNotFoundException: " + e);
            Environment.Exit(-1);
            return null;

        }
        catch (Exception e)
        {
            l.Fatal("!-- EXIT --! at ReadCSV with other Exception: " + e);
            Environment.Exit(-1);
            return null;
        }
    }
    
    public static GraphServiceClient Authentication()
    
    {
        //List<string> scopes = new List<string> { "https://graph.microsoft.com/.default" };
        string tenantId = "56860889-26c2-4d92-8d55-70bc87be7fd8";
        string clientId = "f3ca9eae-5f94-4a1d-8a00-f733cc9e319f";
        string clientSecret = "1Z63vR_.ulHxfk3.YhlL1-lZE6bMFYU_4U";

        IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(clientId)
                                                                                 .WithClientSecret(clientSecret)
                                                                                 .WithTenantId(tenantId)
                                                                                 .Build();
        ClientCredentialProvider authProvider = new ClientCredentialProvider(app);
        GraphServiceClient gc = new GraphServiceClient(authProvider);
        return gc;
    }

    public static List<Feiertag> GetEvents(GraphServiceClient gc, string userId, string today)
    {
        List<IUserCalendarViewCollectionPage> allEvents = new List<IUserCalendarViewCollectionPage>();

        var queryOptions = new List<QueryOption>()
        {
            new QueryOption("startDateTime", today),
            new QueryOption("endDateTime", "2023-12-31T00:00:00.0000000")
        };

        var events = gc.Users[$"{userId}"].CalendarView
                                             .Request(queryOptions)
                                             .Select("subject,bodyPreview,seriesMasterId,type,recurrence,start,end")
                                             .GetAsync()
                                             .Result;
         
 /*       var events = gc.Users[$"{userId}"].Events
                                          .Request(queryOptions)
                                          .Header("Prefer", "outlook.timezone=\"W. Europe Standard Time\"")
                                          .Select("subject,start,end")
                                          .GetAsync()
                                          .Result; */
        

        while (events.Count > 0)
        {
            allEvents.Add(events);
            if (events.NextPageRequest != null)
            {
                events = events.NextPageRequest
                               .GetAsync()
                               .Result;
            }
            else { break;}
        }
        List<Feiertag> allEventsList = new List<Feiertag>();
        
        foreach (IUserCalendarViewCollectionPage page in allEvents)
        {
            foreach (Event event_ in page.CurrentPage)
            {
                Feiertag tempEvent = new Feiertag
                {
                    Subject = event_.Subject,
                    StartDatetime = event_.Start.DateTime.ToString(),
                    EndDateTime = event_.End.DateTime.ToString()
                };
                allEventsList.Add(tempEvent);
            }
        }
        return allEventsList;
    }

    public static Boolean Compare(Feiertag feiertagCsv, List<Functions.Feiertag> feiertageUser)
    //compares subject and startDateTime of current event with all existing events from user. 
    {

        Boolean existing = false;
        foreach (Feiertag feiertag in feiertageUser)
        {
            if (feiertagCsv.Subject == feiertag.Subject & Convert.ToDateTime(feiertagCsv.StartDatetime) == Convert.ToDateTime(feiertag.StartDatetime))
                {
                 //l.Info(feiertag.Subject + " am " + feiertag.StartDatetime + "bereits vorhanden");
                 existing = true;
                 break;
             }
             else { existing = false; } 
        }
        return existing;
    }

    public static Event FormatEvent(Feiertag feiertag)
    {
        //List<string> cat = new List<string> {"Feiertag"};
        var @event = new Event
        {
            ReminderMinutesBeforeStart = 900,
            IsAllDay = true,
            ShowAs = FreeBusyStatus.Oof,
            Categories = new List<string> { "Feiertag" },
            Subject = feiertag.Subject,
            Start = new DateTimeTimeZone
            {
                DateTime = feiertag.StartDatetime,
                TimeZone = "W. Europe Standard Time"
            },
            End = new DateTimeTimeZone
            {
                DateTime = feiertag.EndDateTime,
                TimeZone = "W. Europe Standard Time"
            },
            Location = new Location
            {
                DisplayName = "Österreich"
            }
        };
        return @event;
    }

    public async static Task<Event> PostNewEvent(GraphServiceClient gc, string userId, Event @event)
    //post Event via HTTPS call (RestSharp).
    {
        //logger.Info("Event wird eingetragen...");
        Event t = await gc.Users[$"{userId}"].Events
                                   .Request()
                                   .Header("Prefer", "outlook.timezone=\"W. Europe Standard Time\"")
                                   .AddAsync(@event);

        //logger.Info("Async erfolgreich!");
        return t;
    }
}
