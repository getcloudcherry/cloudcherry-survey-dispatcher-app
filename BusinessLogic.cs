using CloudCherry;
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
using System.Text;
using System.Threading.Tasks;

namespace DispatcherScheduler
{
    class DispatcherHelper
    {
        ScheduledService objservice = new ScheduledService();
        //Retreving CSV string from  CSV or Excel file in ftp or local folder

        public string GetCSVString(DispatcherConfig item)
        {
            objservice.TraceService("Retreiving CSV string");
            string CSVstring = "";

            //If input folder is ftp  then downloading the input file and storing in local path.
            if (item.InputLocationtype.ToLower() == "ftp")
            {
                string localpath = System.IO.Directory.GetCurrentDirectory();
                localpath = localpath.Replace(@"bin\Debug", "Files");


                using (System.Net.WebClient client = new System.Net.WebClient())
                {
                    client.Credentials = new System.Net.NetworkCredential(item.ftpusername, item.ftppassword);
                    string localfile = localpath + @"\survey.xlsx";
                    if (item.InputFileType.ToUpper() == "CSV")
                        localfile = localpath + @"\survey.csv";
                    client.DownloadFile(new Uri(item.InputSource), localfile);
                    item.InputSource = localfile;
                }

            }

            if (item.InputFileType.ToUpper() == "CSV")
            {
                //Retreiving Prefill Data from CSV file
                CSVstring = GetCSVData(item.InputSource);

            }
            else if (item.InputFileType.ToUpper() == "EXCEL")
            {
                //Retreiving CSV string from excel file
                CSVstring = GetExcelData(item.InputSource);
            }
            return CSVstring;
        }

        //Function to Retrieve CSV string from csv file
        public string GetCSVData(string path)
        {
            StreamReader oStreamReader = new StreamReader(path);
            return oStreamReader.ReadToEnd();

        }

        //Function to Retrieve CSV string from Excel file
        public string GetExcelData(string excelpath)
        {
            string fileName = excelpath;
            string Excelstring = "";
            string csvstring = "";

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Row row in rows)
                {
                    try
                    {

                        string thisrow = "";

                        for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                        {
                            string cell = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));

                            if (thisrow == "")
                            {

                                thisrow = cell;

                            }
                            else
                            {
                                thisrow = thisrow + "," + cell;

                            }


                        }
                        if (Excelstring == "")
                            Excelstring = thisrow;
                        else
                            Excelstring = Excelstring + System.Environment.NewLine + thisrow;
                        csvstring = Excelstring;

                    }
                    catch { }
                }


            }
            return csvstring;
        }

        //Retreives tha value in the cell of excel
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            string value = "";
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue != null)
            {
                value = cell.CellValue.InnerXml;
            }
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }



        //Constructing a table from the output response of cloud cherry
        DataTable Getsurveytable(string survey)
        {
            string[] ch = new string[1] { Environment.NewLine };
            string[] surveydata = survey.Split(ch, StringSplitOptions.RemoveEmptyEntries);
            string[] qids = surveydata[1].Split(',');
            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn();
            for (int i = 0; i < qids.Length; i++)
            {
                dc = new DataColumn();
                dc.ColumnName = qids[i].Replace(" ", "_");
                dt.Columns.Add(dc);
            }
            for (int i = 2; i < surveydata.Length; i++)
            {
                string[] data = surveydata[i].Split(',');
                DataRow dr = dt.NewRow();
                for (int j = 0; j < data.Length; j++)
                {
                    dr[j] = data[j];
                }
                dt.Rows.Add(dr);
            }


            return dt;
        }

        //retreiving the  column based on tag
        string GetQuestionfortag(string tag, List<Question> Ques)
        {

            foreach (var q in Ques)
            {
                if (q.QuestionTags == null)
                    continue;
                if (q.QuestionTags.Contains(tag))
                {
                    return q.Text.Replace(" ", "_");
                }
            }

            objservice.TraceService("No question tag exist with name :" + tag);
            return "";
        }

        //Deserialising the content in  the Config.json file
        public List<DispatcherConfig> GetDispatcherlist()
        {
            string filePath = ConfigurationManager.AppSettings["JsonConfigPath"];
            List<DispatcherConfig> items = new List<DispatcherConfig>();
            using (StreamReader r = new StreamReader(filePath))
            {
                string json = r.ReadToEnd();
                items = JsonConvert.DeserializeObject<List<DispatcherConfig>>(json);
            }            
            return items;
        }



        //Writing the output in to CSV file in local folder or ftp folder
        public void WriteCSV(DispatcherConfig objdispatcher, string csvstring, bool iscreate)
        {
            string path = objdispatcher.OutputDestination;

            //Creating  temporary local path to store output response if the output location  is ftp folder
            if (objdispatcher.OutputLocationType.ToLower() == "ftp")
            {
                objservice.TraceService("output location type : " + objdispatcher.OutputLocationType);
                string localpath = System.IO.Directory.GetCurrentDirectory();
                path = localpath.Replace(@"bin\Debug", @"Files\");

            }
            path = path + "SurveyTokens_" + objdispatcher.CloudCherryAccount.ToUpper() + ".csv";
            objservice.TraceService("Writing into output csv file:" + path);
            FileStream fw;
            if (iscreate)
                fw = new FileStream(path, FileMode.Create);
            else
            {
                fw = new FileStream(path, FileMode.Append);
                string[] delimiter = { "\r\n" };
                string[] csvlist = csvstring.Split(delimiter, 3, StringSplitOptions.None);
                csvstring = csvlist[2];
            }


            using (StreamWriter sw = new StreamWriter(fw))
            {

                sw.Write(csvstring);
                sw.WriteLine();
            }

            //Uploading the final output file to ftp folder and deleting the temporary output file
            if (objdispatcher.OutputLocationType.ToLower() == "ftp")
            {

                string ftppath = objdispatcher.OutputDestination + "SurveyTokens_" + objdispatcher.CloudCherryAccount.ToUpper() + ".csv";

                System.Threading.AutoResetEvent waiter = new System.Threading.AutoResetEvent(false);

                using (System.Net.WebClient client = new System.Net.WebClient())
                {
                    client.Credentials = new System.Net.NetworkCredential(objdispatcher.ftpusername, objdispatcher.ftppassword);

                    client.UploadFileCompleted += new UploadFileCompletedEventHandler(UploadFileCallback);
                    client.UploadFileAsync(new Uri(ftppath), "STOR", path, waiter);
                    waiter.WaitOne();
                    objservice.TraceService("File uploaded to ftp");
                    File.Delete(path);
                }
            }

        }

        private static void UploadFileCallback(Object sender, UploadFileCompletedEventArgs e)
        {
            System.Threading.AutoResetEvent waiter = (System.Threading.AutoResetEvent)e.UserState; ;
            try
            {
                string reply = System.Text.Encoding.UTF8.GetString(e.Result);
                ScheduledService objservice = new ScheduledService();
                objservice.TraceService(reply);
            }
            finally
            {
                waiter.Set();
            }
        }


        //Validating if the given prefill questions  in input file is valid or not
        public bool ValidatingPrefillQuestions(string CSVstring, List<Question> Ques)
        {
            objservice.TraceService("Validating Questions");
            int qcounter = 0;
            string[] delimiter = { "\r\n" };
            string[] csvlist = CSVstring.Split(delimiter, StringSplitOptions.None);
            delimiter[0] = ",";
            string[] questinids = csvlist[0].Split(delimiter, StringSplitOptions.None);
            foreach (Question q in Ques)
            {
                if (q.StaffFill)
                {
                    if (questinids.Contains(q.Id))
                    { qcounter++; }
                }
            }
            return qcounter == questinids.Length;
        }

        //Splitting CSV string  for every 50k rows  to upload bulk tokens
        public List<string> SplitsCSVstring(string csvstring)
        {
            objservice.TraceService("Splitting csv string to bulk upload");
            string[] delimiter = { "\r\n" };
            List<string> finalcsv = new List<string>();
            string[] csvlist = csvstring.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);
            int csvsplitvalue = 50000;
            for (int i = 2; i < csvlist.Length; i++)
            {
                int index = i / csvsplitvalue;
                if (index + 1 > finalcsv.Count)
                {
                    finalcsv.Add("");
                    finalcsv[index] = csvlist[0] + "\r\n" + csvlist[1];
                }
                finalcsv[index] = finalcsv[index] + "\r\n" + csvlist[i];
            }

            return finalcsv;
        }

        //Validating the input types mentioned in Config.json file
        public bool ValidatingFileTypes(DispatcherConfig item)
        {

            if (item.OutputLocationType.ToLower() == "ftp")
            {
                if (item.OutputType.ToLower() == "email" || item.OutputType.ToLower() == "sms")
                {
                    objservice.TraceService("output type cannot be " + item.OutputType + " when output location type is ftp");
                    return false;
                }
            }

            if ((item.InputFileType.ToLower() == "csv") && (Path.GetExtension(item.InputSource).ToLower() != ".csv"))
            {
                objservice.TraceService("Input file Extension '" + Path.GetExtension(item.InputSource) + "' is not matching with input file type :'" + item.InputFileType + "'");

                return false;

            }
            if ((item.InputFileType.ToLower() == "excel") && (Path.GetExtension(item.InputSource).ToLower() != ".xlsx"))
            {
                objservice.TraceService("Input file Extension '" + Path.GetExtension(item.InputSource) + "' is not matching with input file type :'" + item.InputFileType + "'");

                return false;

            }

            if ((item.OutputType.ToLower() == "csv") && (Path.GetExtension(item.OutputDestination).ToLower() == ".csv"))
            {
                objservice.TraceService("output file Extension  '" + Path.GetExtension(item.OutputDestination) + "' is not matching with input file type :'" + item.OutputType + "'");

                return false;
            }

            if ((item.InputFileType.ToLower() == "excel") || (item.InputFileType.ToLower() == "excel"))
            {
                if (!File.Exists(item.InputSource))
                {
                    objservice.TraceService("Input file  does not exist :" + item.InputSource);
                    return false;
                }
            }

            if (item.OutputType.ToLower() == "csv")
            {
                if (!Directory.Exists(item.OutputDestination))
                {
                    objservice.TraceService("Output file  does not exist :" + item.InputSource);
                    return false;
                }
            }

            return true;
        }

        //Sending Sms to the numbers mentioned in input file
        public void SendSMS(List<Question> Ques, string outputresponse, DispatcherConfig objDispatcher)
        {
            string outputsource = objDispatcher.OutputDestination;
            string namecolumn = GetQuestionfortag("firstname", Ques);
            if (namecolumn == "")
                return;

            //Constructing a table from the output response of cloud cherry
            DataTable dt = Getsurveytable(outputresponse);

            //retreiving the mobile column
            string column = GetQuestionfortag("Mobile", Ques);
            if (column == "")
                return;
            for (int r = 0; r < dt.Rows.Count; r++)
            {

                string mobile = dt.Rows[r][column].ToString();
                string url = dt.Rows[r]["Survey_URL"].ToString();

                string name = dt.Rows[r][namecolumn].ToString();
                string parameters = objDispatcher.OutputDestination;
                parameters = parameters.Replace("$msg", objDispatcher.Salutation + "\n" + objDispatcher.Message + "\n" + objDispatcher.Signature);
                parameters = parameters.Replace("$url", url);
                parameters = parameters.Replace("$to", mobile);
                objservice.TraceService("sending sms to " + mobile + " with url " + url + " (" + (r + 1) + " of " + dt.Rows.Count + ")");
                var request = (HttpWebRequest)WebRequest.Create(parameters);//sending sms
                request.GetResponse();

                System.Threading.Thread.Sleep(objDispatcher.OutputDelay);
            }
        }
        //Retrieving respective email and mailing the token url details
        public void SendEmail(List<Question> Ques, string outputresponse, DispatcherConfig objDispatcher)
        {

            string namecolumn = GetQuestionfortag("firstname", Ques);

    if (namecolumn == "")
                return;            //Constructing a table from the output response of cloud cherry
            DataTable dt = Getsurveytable(outputresponse);

            ///retreiving the email column
            string column = GetQuestionfortag("Email", Ques);
 if (column == "")
                return;
            for (int r = 0; r < dt.Rows.Count; r++)
            {

                string email = dt.Rows[r][column].ToString();
                string url = dt.Rows[r]["Survey_URL"].ToString();
                string name = dt.Rows[r][namecolumn].ToString();
                string parameters = objDispatcher.OutputDestination;
                parameters = parameters.Replace("$msg", objDispatcher.Salutation + "\n" + objDispatcher.Message + "\n" + objDispatcher.Signature);
                parameters = parameters.Replace("$url", url);
                parameters = parameters.Replace("$to", email);

                objservice.TraceService("sending mail to " + email + " with url " + url + " (" + (r + 1) + " of " + dt.Rows.Count + ")");
                SendEmail(parameters);//sending mails
                System.Threading.Thread.Sleep(objDispatcher.OutputDelay);
            }
        }

        //Sending emails to the mailids mentioned in input file
        public void SendEmail(string parameters)
        {
            try
            {

                HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create(parameters);
                myHttpWebRequest.Method = "POST";
                myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";
                ServicePointManager.DefaultConnectionLimit = 100;
                using (var http = myHttpWebRequest.GetRequestStream())
                {
                    StreamWriter streamWriter = new StreamWriter(http);
                    streamWriter.Write(parameters);
                    streamWriter.Flush();
                    streamWriter.Close();
                    HttpWebResponse httpResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                }
            }
            catch (WebException ex)
            {
                objservice.TraceService(ex.Message);

            }


        }
    }

    public class DispatcherConfig
    {
        //Account 
        public string CloudCherryAPIEndPoint { get; set; } // defaults to https://api.getcloudcherry.com
        public string CloudCherryAccount { get; set; } // username
        public string CloudCherrySecret { get; set; } // password

        public string InputLocationtype { get; set; } // InputLocationtype
        public string ftpusername { get; set; } // ftpusername
        public string ftppassword { get; set; } // ftppassword
        public string OutputLocationType { get; set; } // OutLocationtype
        //Survey
        public string SurveyLocation { get; set; } // "Downtown"
        public int SurveyValidFor { get; set; } // Days(Max is 90)
        public int SurveyUses { get; set; }

        public string InputFileType { get; set; } // CSV/Excel/ODBC(Win Table)
        public string InputSource { get; set; } // Filename.csv/Filename.xlsx/ConnectionString(ODBC)
        internal Dictionary<string, string> InputMapping { get; set; } // CSV/Excel/Table Column to QuestionID
        public string Salutation { get; set; }
        public string Message { get; set; }
        public string Signature { get; set; }
        public string OutputType { get; set; } // CSV, URL(SMS), SMTP(Email), ODBC Update(Table)
        public int OutputDelay { get; set; } // Millisecond delay between calls for SMS/Email to enable not overloading with millions of emails/sms in one go
        public string OutputDestination { get; set; } // https://x.y.z?sms=$NUM , "outfile.csv", "username:password@smtp.xyz.com:587/sender name/sender address"
    }
}
