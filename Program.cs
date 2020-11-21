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
            //set todays date to get only future events.
            string today = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");

            //UserID müsst ma noch wo anders herbekommen... <---------------------------------------!!!!
            string userId = "test@sebastiansoldan.onmicrosoft.com";

            //info message
            l.Info($"Starting with userID: {userId} (today = {today})"); 
            l.Info("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^");

            //load events from csv-file to an List of "Feiertag"-Objects.
            List<Functions.Feiertag> feiertageFromCsv = Functions.ReadCSV(@"C:\Users\sebas\source\repos\FeiertagAppV0.2.0\Feiertage.csv", today);

            //start authentication
            GraphServiceClient gc = Functions.Authentication();

            //get all future Events from User via GraphAPI
            List<Functions.Feiertag> existingEventsFromUser = Functions.GetUserEvents(gc, userId, today).Result;
            
            int duplicates = 0; //for counting duplicates.. (DEBUG)   
            int batchId = 0; //needed for "batch-posting" events

            l.Info("--> Starting to compare...");

            BatchRequestContent batch = new BatchRequestContent(); //batch is used for posting 20 events with on http-call.

            try
            {
                int numbersOfBatches = 0;
                foreach (Functions.Feiertag feiertag in feiertageFromCsv)
                {
                    //compare with UserEvents if the event is allready existing.
                    bool existing = Functions.Compare(feiertag, existingEventsFromUser);
                    if (existing) //do not post event.
                    {
                        l.Info("---->" + feiertag.Subject + " on " + feiertag.StartDatetime + " allready existing!");
                        duplicates++;
                    }
                    else //post event via batch.
                    {
                        l.Info($"------> {feiertag.Subject} on {feiertag.StartDatetime} added to batch {numbersOfBatches}.");
                        Event @event = Functions.FormatEvent(feiertag); //format data to json

                        if (batchId == 0) 
                    //first added BatchRequestStep does not need "dependingOn"
                        {

                            batch.AddBatchRequestStep(new BatchRequestStep(batchId.ToString(), Functions.BuildHttpMessage(gc, userId, @event)));
                            batchId++;
                        }

                        else if (batchId > 0 && batchId < 20)
                    //if its an following BatchRequestStep "dependingOn" = batchId from BatchRequestSteps before.
                        {
                            batch.AddBatchRequestStep(new BatchRequestStep(batchId.ToString(), Functions.BuildHttpMessage(gc, userId, @event), new List<string> { (batchId - 1).ToString() }));
                            batchId++;
                        }

                        if (batchId == 20)
                    //Batch has an limit of 20 BatchRequestSteps. If this is reached, batch gets posted and all BatchRequestSteps are going to be removed so they can be refilled in next foreach rounds.
                        {
                            numbersOfBatches++;
                            batch.AddBatchRequestStep(new BatchRequestStep(batchId.ToString(), Functions.BuildHttpMessage(gc, userId, @event), new List<string> { (batchId - 1).ToString() }));
                            var task1 = gc.Batch.Request().PostAsync(batch);
                            task1.Wait();
                            int j = 20;
                            while (batch.BatchRequestSteps.Count > 0)
                            {
                                batch.RemoveBatchRequestStepWithId(j.ToString());
                                j--;
                            }
                            batchId = 0;
                        }
                    }
                }

                if (batch.BatchRequestSteps.Count != 0) //makes sure that batch is postet, when there are BatchRequestSteps left after the foreach Rounds above.
                {
                    numbersOfBatches++;
                    var task = gc.Batch.Request().PostAsync(batch);
                    task.Wait();
                }

                //info message
                l.Info($"{feiertageFromCsv.Count - duplicates} Feiertage from {feiertageFromCsv.Count} postet in {numbersOfBatches} batches. {duplicates} allready existed.");
                l.Info("---------------------------------------------------------------------------------------\n");
            }              
            catch (Exception e)
            {
                l.Error("!-- FAILED --! at Batching with Exception: " + e);
                Environment.Exit(-9);
            }
        }
    }
}
