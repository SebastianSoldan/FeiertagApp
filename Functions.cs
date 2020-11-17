using NLog;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System.Text.Json;
using System.Text;

public static class Functions
{
    private static NLog.Logger l = LogManager.GetCurrentClassLogger();

    public class Feiertag
    {
        public string Subject { get; set; }
        public string StartDatetime { get; set; }
        public string EndDateTime { get; set; }
    }

    public class HttpClientFactory : IMsalHttpClientFactory
    //needed for proxy use.
    {
        public HttpClient GetHttpClient()
        {
            var proxy = new WebProxy
            {
                Address = new Uri(System.Configuration.ConfigurationManager.AppSettings["proxyurl"]),
                UseDefaultCredentials = true
            };
            HttpClientHandler hch = new HttpClientHandler()
            {
                Proxy = proxy,
                UseProxy = true
            };

            return new HttpClient(hch);
        }
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
                if (DateTime.Compare(Convert.ToDateTime(tempFeiertag.StartDatetime), Convert.ToDateTime(today)) >= 0)
                //make sure to add future-events only
                {
                    FeiertageAusCsv.Add(tempFeiertag);
                }
            }
            return FeiertageAusCsv;
        }
        catch (FileNotFoundException e)
        {
            l.Error("!-- EXIT --! at ReadCSV with FileNotFoundException: " + e);
            Environment.Exit(-1);
            return null;
        }
        catch (Exception e)
        {
            l.Error("!-- EXIT --! at ReadCSV with other Exception: " + e);
            Environment.Exit(-1);
            return null;
        }
    }

    public static GraphServiceClient Authentication()
    //handles authentication.
    {
        //should come from web.config
        string tenantId = "56860889-26c2-4d92-8d55-70bc87be7fd8";
        string clientId = "f3ca9eae-5f94-4a1d-8a00-f733cc9e319f";
        string clientSecret = "1Z63vR_.ulHxfk3.YhlL1-lZE6bMFYU_4U";

        //create a httpclienthandler for proxy use
 /*-------------------------------------------------------------------------------------------------------------------
        var proxy = new WebProxy
        {
            Address = new Uri(System.Configuration.ConfigurationManager.AppSettings["proxyurl"]),
            UseDefaultCredentials = true
        };
        HttpClientHandler hch = new HttpClientHandler
        {
            Proxy = proxy,
            UseProxy = true
        };
        var hcf = new HttpClientFactory();
        var httpProvider = new HttpProvider(hch, false); 
 ----------------------------------------------------------------------------------------------------------------------*/

        //build confidential client app - this is needed for MSAL.NET to gain authentication
        IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(clientId)
                                                                                 .WithClientSecret(clientSecret)
                                                                                 .WithTenantId(tenantId)
                                                                                 //.WithHttpClientFactory(hcf)
                                                                                 .Build();

        //create an AuthProvider from Application. This is used by GraphServiceClient
        ClientCredentialProvider authProvider = new ClientCredentialProvider(app);
        GraphServiceClient gc = new GraphServiceClient(authProvider/*, httpProvider*/);

        return gc;
    }


    public static async Task<List<Feiertag>> GetUserEvents(GraphServiceClient gc, string userId, string today)
    {
        //creat list to store pages from response
        List<IUserCalendarViewCollectionPage> pages = new List<IUserCalendarViewCollectionPage>();


        //set the timeframe for Events wich are going to be called.
        var queryOptions = new List<QueryOption>()
        {
            new QueryOption("startDateTime", today),
            new QueryOption("endDateTime", "2023-12-30T00:00:00.00")
        };

        //build Request to get existing events from users calender.
        var response = await gc.Users[$"{userId}"].CalendarView
                                                  .Request(queryOptions)
                                                  .Select("subject,start,end")
                                                  .GetAsync();

        /*var response = await gc.Users[$"{userId}"].Events
                                                 .Request()
                                                 .Header("Prefer", "outlook.timezone=\"W. Europe Standard Time\"")
                                                 .Select("subject,start,end")
                                                 .GetAsync();*/

        //request next page until all events are loaded
        pages.Add(response);
        while (response.NextPageRequest != null)
        {
            response = await response.NextPageRequest
                                     .GetAsync();

            pages.Add(response);
        } // ---------------------------------------------------------------------------!!!!!!! vereinfachen
        
        //create list to store events in easy format.
        List<Feiertag> allEvents = new List<Feiertag>();
        
        //add events from pages to list
        foreach (IUserCalendarViewCollectionPage page in pages)
        {
            foreach (Event event_ in page.CurrentPage)
            {
                Feiertag tempEvent = new Feiertag
                {
                    Subject = event_.Subject,
                    StartDatetime = event_.Start.DateTime.ToString(),
                    EndDateTime = event_.End.DateTime.ToString()
                };
                allEvents.Add(tempEvent);
            }
        }

        return allEvents;
    }

    public static bool Compare(Feiertag feiertagCsv, List<Feiertag> feiertageUser)
    //compares subject and startDateTime of current event with all existing events from user. 
    {
        Boolean existing = false;
        foreach (Feiertag feiertag in feiertageUser)
        {
            if (feiertagCsv.Subject == feiertag.Subject && Convert.ToDateTime(feiertagCsv.StartDatetime) == Convert.ToDateTime(feiertag.StartDatetime))
                {
                 existing = true;
                 break;
                }
            else 
            {
                existing = false; 
            }
        }
        return existing;
    }

    public static Event FormatEvent(Feiertag feiertag)
    //format data to JSON (needed for postEventAsync)
    {
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

    public static HttpRequestMessage BuildHttpMessage(GraphServiceClient gc, string userId, Event @event)
    //post Event async with Graph Service Client.
    {
        var ser = new Serializer();

        var t = gc.Users[$"{userId}"].Events
                           .Request()
                           .GetHttpRequestMessage();

        t.Method = HttpMethod.Post;
        t.Content = ser.SerializeAsJsonContent(@event);
        t.Headers.Add("Prefer", "outlook.timezone=\"W. Europe Standard Time\"");
    
        return t;
    }
}
