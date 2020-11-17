using NLog;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Graph;

namespace FeiertagAppV0._2._0
{
    class Program
    {
        private static Logger l = LogManager.GetCurrentClassLogger();

        static void Main()
        {
            string today = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");

            List<Functions.Feiertag> feiertageFromCsv = Functions.ReadCSV(@"C:\Users\sebas\source\repos\FeiertagAppV0.2.0\Feiertage.csv", today);

            string userId = "test@sebastiansoldan.onmicrosoft.com";

            GraphServiceClient gc = Functions.Authentication();

            List<Functions.Feiertag> exinstingEventsFromUser = Functions.GetEvents(gc, userId, today);

            int duplicates = 0, success = 0;
            var tasks = new List<Task>();
            foreach (Functions.Feiertag feiertag in feiertageFromCsv)
            {
                //Compare with UserEvents if the event is allready existing.
                Boolean existing = Functions.Compare(feiertag, exinstingEventsFromUser);
                if (existing) //do not post event.
                {
                    l.Info(feiertag.Subject + " am " + feiertag.StartDatetime + " bereits vorhanden!");
                    duplicates++;
                }
                else //post event.
                {
                    l.Info(feiertag.Subject + " am " + feiertag.StartDatetime + " GEPOSTET");
                    Event @event = Functions.FormatEvent(feiertag);
                    Task<Event> t = Functions.PostNewEvent(gc, userId, @event);
                    tasks.Add(t);
                    Thread.Sleep(180);
                    success++;
                }
            }
            Task.WaitAll(tasks.ToArray());
            //write summary to Console.
            //logger.Info($"Es wurden {success} von {feiertageFromCSV.Count} Feiertagen erfolgreich eingetragen. {duplicates} waren bereits vorhanden.");
            //logger.Info("-----------------------------------------------------------------------------------------------");


            Console.WriteLine("BLAABLAABLAA");
        }
    }
}
