using NLog;
using System;
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
            //string today = "2022-12-30T00:00:00.00";
            string userId = "test@sebastiansoldan.onmicrosoft.com";
            l.Info($"-- Start Programm for {userId} startTime = {today}");

            List<Functions.Feiertag> feiertageFromCsv = Functions.ReadCSV(@"C:\Users\sebas\source\repos\FeiertagAppV0.2.0\Feiertage.csv", today);

            GraphServiceClient gc = Functions.Authentication();

            List<Functions.Feiertag> existingEventsFromUser = Functions.GetUserEvents(gc, userId, today).Result;

            int duplicates = 0;
            int batchId = 0;
            //var tasks = new List<Task>();
            BatchRequestContent batch = new BatchRequestContent();
            foreach (Functions.Feiertag feiertag in feiertageFromCsv)
            {
                //Compare with UserEvents if the event is allready existing.
                bool existing = Functions.Compare(feiertag, existingEventsFromUser);
                if (existing) //do not post event.
                { 
                    l.Info(feiertag.Subject + " am " + feiertag.StartDatetime + " bereits vorhanden!");
                    duplicates++;
                }
                else //post event.
                {
                    l.Info("---> " + feiertag.Subject + " am " + feiertag.StartDatetime + " posten.");
                    Event @event = Functions.FormatEvent(feiertag);
                    if (batchId == 0)
                    {
                        
                        batch.AddBatchRequestStep(new BatchRequestStep(batchId.ToString(), Functions.BuildHttpMessage(gc, userId, @event)));
                        batchId++;
                    }
                        
                    else if (batchId > 0 && batchId < 20)
                    {
                        //batch.AddBatchRequestStep(Functions.BuildHttpMessage(gc, userId, @event)); //---------------------!!!!!
                        batch.AddBatchRequestStep(new BatchRequestStep(batchId.ToString(), Functions.BuildHttpMessage(gc, userId, @event), new List<string> { (batchId - 1).ToString() }));
                        batchId++;
                    }
                    if (batchId == 20)
                    {
                        batch.AddBatchRequestStep(new BatchRequestStep(batchId.ToString(), Functions.BuildHttpMessage(gc, userId, @event), new List<string> { (batchId - 1).ToString() }));
                        var task1 = gc.Batch.Request().PostAsync(batch);
                        task1.Wait();
                        int j = 20;
                        while(batch.BatchRequestSteps.Count > 0)
                        {
                            batch.RemoveBatchRequestStepWithId(j.ToString());
                            j--;
                        }

                        //------------------------exception
                        batchId = 0;
                    }
                }
            }
            var task = gc.Batch.Request().PostAsync(batch);
            task.Wait();

            //var task = gc.Batch.Request().PostAsync(batch);

            l.Info($"Es wurde versucht {feiertageFromCsv.Count} Feiertage einzutragen. {duplicates} waren bereits vorhanden.");
            l.Info("-----------------------------------------------------------------------------------------------");
        }
    }
}
