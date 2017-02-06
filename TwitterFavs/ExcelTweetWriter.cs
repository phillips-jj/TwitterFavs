using System;
using Microsoft.Office.Interop.Excel;

namespace TwitterFavs
{
    class ExcelTweetWriter
    {
        bool mSane = false;
        private Microsoft.Office.Interop.Excel.Application mApp = null;
        private Workbook mWB = null;
        private Worksheet mWS = null;
        private string mFilename;
        private long mNextAvailableRow = 2;

        private ExcelTweetWriter() { }  //no default constructor

        public ExcelTweetWriter(string sFilename)
        {
            mFilename = sFilename;
            mApp = new Application();
            mApp.DisplayAlerts = false;         //prevents the SaveAs dialog from popping up an overwrite confirmation if the file already exists
            mWB = mApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            mWS = (Worksheet)mWB.Worksheets[1];

            //set up column headers
            mWS.Range["A1"].Value2 = "ID";
            mWS.Range["B1"].Value2 = "Date";
            mWS.Range["C1"].Value2 = "User Name";
            mWS.Range["D1"].Value2 = "@ Name";
            mWS.Range["E1"].Value2 = "Tweet";

            //future enhancement - if there is a photo in the tweet, download it an insert it into column F

            Range r = mWS.Range["A1", "E1"];
            r.Font.Bold = true;
            r.Font.Underline = true;

            mSane = true;
        }

        ~ExcelTweetWriter()
        {
            Close();
        }

        public void Close()
        {
            mSane = false;

            if (mWB != null)
            {
                //make sure the columns are wide enough to see the entire text
                Range r = (Range)mWS.Columns["A:E"];
                r.AutoFit();

                try
                {
                    //warning - because we set mApp.DisplayAlerts=false above, if the file already exists, it will be overwritten silently
                    mWB.SaveAs(mFilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   XlSaveAsAccessMode.xlExclusive, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing);
                    mWB.Close();
                }
                catch(Exception e)
                {
                    Type t = e.GetType();

                }

                mWB = null;
            }

            if (mApp != null)
            {
                mApp.Quit();
                mApp = null;
            }
        }

        public bool Write(DateTime created, long ID, string atName, string displayName, string tweetText)
        {
            if (!mSane) return false;

            Range r = mWS.Range["A" + mNextAvailableRow];
            r.Value2 = ID;
            r.NumberFormat = "0";   //display the numbers as-is, not in exponential format

            //present date/time values as ISO 8601-ish format
            r = mWS.Range["B" + mNextAvailableRow];
            r.Value2 = created;
            r.NumberFormat = "YYYY-MM-DD hh:mm UTC";

            
            mWS.Range["C" + mNextAvailableRow].Value2 = atName;
            mWS.Range["D" + mNextAvailableRow].Value2 = displayName;
            mWS.Range["E" + mNextAvailableRow].Value2 = tweetText;

            mNextAvailableRow++;
            return true;
        }
    }
}
