using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;

namespace FinDataAlphaVantage
{
    internal class FinDataSet
    {
        internal const string MyApiKey = "";
        internal readonly string ApiKey = "";
        internal readonly string Symbol = "";
        internal readonly TimeSeries Series = TimeSeries.None;

        internal bool RetrieveRawDataSuccessful { get; private set; } = false;
        internal string RawData { get; private set; } = "";

        internal bool ProcessDataSuccessful { get; private set; } = false;
        internal List<string> LineByLineData { get; private set; } = new();



        internal FinDataSet(string apiKey, string symbol, TimeSeries series)
        {
            ApiKey = apiKey;
            Symbol = symbol;
            Series = series;
        }


        internal void GetRawDataFromCSV(string path)
        {
            RawData = File.ReadAllText(path);
            RetrieveRawDataSuccessful = true;
        }


        internal async Task GetRawDataFromAlphaVantage()
        {
            string url = $"https://www.alphavantage.co/query?function={ForURLString(Series)}&symbol={Symbol}&apikey={ApiKey}&datatype=csv";

            using HttpClient client = new HttpClient();
            try
            {
                RawData = await client.GetStringAsync(url);
                RetrieveRawDataSuccessful = true;
            }
            catch(Exception exception)
            {
                MessageBox.Show(exception.Message, "Error");
            }
        }

        internal void WriteRawDataToFile(string completeFilePath)
        {
            Debug.Assert(RetrieveRawDataSuccessful, "Cannot print data. Retrieve was not successful (yet).");
            File.WriteAllText(completeFilePath, RawData);
        }


        private static string ForURLString(TimeSeries timeSeries)
        {
            switch ( timeSeries )
            {
                case TimeSeries.None:
                    return "";
                case TimeSeries.Intraday:
                    return "TIME_SERIES_INTRADAY";
                case TimeSeries.Daily:
                    return "TIME_SERIES_DAILY";
                case TimeSeries.DailyAdjusted:
                    return "TIME_SERIES_DAILY_ADJUSTED";
                case TimeSeries.Weekly:
                    return "TIME_SERIES_WEEKLY";
                case TimeSeries.WeeklyAdjusted:
                    return "TIME_SERIES_WEEKLY_ADJUSTED";
                case TimeSeries.Monthly:
                    return "TIME_SERIES_MONTHLY";
                case TimeSeries.MonthlyAdjusted:
                    return "TIME_SERIES_MONTHLY_ADJUSTED";
                default:
                    return "";

            }

        }

    }
   
    internal enum TimeSeries { None, Intraday, Daily, DailyAdjusted, Weekly, WeeklyAdjusted, Monthly, MonthlyAdjusted };


}
