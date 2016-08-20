using CloudCherry;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
namespace DispatcherScheduler
{
    class CloudCherry
    {
        ScheduledService objservice = new ScheduledService();
        DispatcherHelper objhelper = new DispatcherHelper();

        public async void ImportData()
        {
            DispatcherConfig objDispatcher = new DispatcherConfig();
            try
            {
                //Recording log 
                objservice.TraceService("Deserializing ");

                List<DispatcherConfig> objDispatcherlist = objhelper.GetDispatcherlist();//Deserialising the content in  the Config.json file               
                string CSVstring = "";
                objDispatcher = objDispatcherlist[0];
                if (!File.Exists(objDispatcher.InputSource))
                {
                    File.Delete(ScheduledService.schedularpath);
                    return;
                }
                if (!objhelper.ValidatingFileTypes(objDispatcher))
                {
                    return;

                }
                //List of comma separated strings which holds data that needs to be bulk uploaded using survey token
                List<string> CSVstringarr = new List<string>();

                //Initiating CloudCherry API Endpoint with CloudCherryAPIEndPoint, CloudCherryAccount, CloudCherrySecret Specified in config.json
                APIClient client = new APIClient(objDispatcher.CloudCherryAPIEndPoint, objDispatcher.CloudCherryAccount, objDispatcher.CloudCherrySecret);

                //Authenticating user using CloudCherry API
                if (!await client.Login())
                {
                    //CloudCherry Authentication failed writing to Log
                    objservice.TraceService("Login Failed");
                    return;
                }

                //Retreiving all the active questions for the user for validating with the questions given the Input(Excel/CSV)
                List<Question> Ques = await client.GetQuestions(true);

                //Writing in to the log what Input type was selected for bulk import of tokens
                objservice.TraceService("Input Type :" + objDispatcher.InputFileType);

                //Retreiving comma separated string of bu
                CSVstring = objhelper.GetCSVString(objDispatcher);

                //Validating the prefill questions  in input files(excel or csv)
                if (!objhelper.ValidatingPrefillQuestions(CSVstring, Ques))
                {
                    objservice.TraceService("Invalid questions encountered");
                    return;
                }


                CSVstringarr = objhelper.SplitsCSVstring(CSVstring);
                if (CSVstringarr.Count == 0)
                {
                    objservice.TraceService("No answers found");
                    return;
                }

                for (int a = 0; a < CSVstringarr.Count; a++)
                {

                    objservice.TraceService("Uploading Bulk tokens " + (a + 1) + " of " + CSVstringarr.Count);
                    //Bulk Uploading csv string with number of rows <=50000 at a time
                    string Outputresponse = await client.UploadBulkTokens(CSVstringarr[a], objDispatcher.SurveyValidFor, objDispatcher.SurveyUses, objDispatcher.SurveyLocation);
                    if (Outputresponse == null)
                    {
                        objservice.TraceService("Invalid data encountered");
                        return;
                    }
                    objservice.TraceService("Output Type :" + objDispatcher.OutputType);
                    if (objDispatcher.OutputType == "CSV")
                    {//Writing response to CSV file in local folder or ftp folder
                        objhelper.WriteCSV(objDispatcher, Outputresponse, a == 0);
                    }

                    if (objDispatcher.OutputType == "Email")
                    {//Sending  emails
                        objhelper.SendEmail(Ques, Outputresponse, objDispatcher);

                    }
                    if (objDispatcher.OutputType == "SMS")
                    {//Sending smses
                        objhelper.SendSMS(Ques, Outputresponse, objDispatcher);
                    }
                }


            }
            catch (Exception ee)
            {
                objservice.TraceService(ee.Message);
            }
            finally
            {
              
                if (File.Exists(objDispatcher.InputSource))
                {
                    objservice.TraceService("completed");
                    string CompletedFilePath = ConfigurationManager.AppSettings["CompletedFilePath"];
                    CompletedFilePath = CompletedFilePath + DateTime.Now.Year + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + @"\";
                    Directory.CreateDirectory(CompletedFilePath);
                    CompletedFilePath = CompletedFilePath + Path.GetFileName(objDispatcher.InputSource);
                    File.Move(objDispatcher.InputSource, CompletedFilePath);
                    File.Delete(objDispatcher.InputSource);
                }
            }
        }
    }
}
