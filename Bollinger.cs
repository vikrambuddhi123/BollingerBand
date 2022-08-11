using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;



namespace BollingerBand
{
    public static class CustomExtensions
    {
        public static IEnumerable<double?> MovingAverage(this IEnumerable<double> source, int windowSize)
        {
            var queue = new Queue<double>(windowSize);
            foreach (double d in source)
            {
                if (queue.Count == windowSize)
                {
                    queue.Dequeue();
                }
                queue.Enqueue(d);
                if (queue.Count == windowSize)
                    yield return queue.Average();
                else
                    yield return null;
            }
        }

        public static IEnumerable<double?> StandardDeviation(this IEnumerable<double> source, int windowSize)
        {
            var queue = new Queue<double>(windowSize);
            foreach (double d in source)
            {
                if (queue.Count == windowSize)
                {
                    queue.Dequeue();
                }
                queue.Enqueue(d);
                double avg = queue.Average();
                double sum = queue.Sum(dd => Math.Pow(dd - avg, 2));
                if (queue.Count == windowSize)
                    yield return Math.Sqrt((sum) / (queue.Count() - 1));
                else
                    yield return null;
            }
        }
    }


    public class Bollinger
    {      
        
        private int N = 20;  // window size for moving average and standard deviation 
        private int K = 2;   // Bollinger Band multple of standard deviation

        public string intraRunDate;
        Dictionary<string, double> rundaySMA,rundayLBB,rundayUBB;

        public DataTable eodData, smaData, stdData, lowerBBData, upperBBData, intraData, orders;
        public Dictionary<string, int> orderPosition;
        public Dictionary<string, double> orderProfit, orderEntryPrice;


        public Bollinger(string historicalEODfile)
        {
            
            // read historical data and compute SMA(20), STDV(20), and Bollinger Bands
            eodData = new DataTable();
            eodData = Utility.ReadExcel(@"..\..\" + historicalEODfile);
            eodData.Columns["Column1"].ColumnName = "Date";

            smaData = eodData.Copy();
            stdData = eodData.Copy();
            lowerBBData = eodData.Copy();
            upperBBData = eodData.Copy();


            // calculate the sma, stdv, bollinger bands
            var cols = eodData.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            foreach (var colname in cols.GetRange(1, 6))
            {
                var series = eodData.AsEnumerable().Select(r => r.Field<double>(colname)).ToList();
                List<double?> sma = series.MovingAverage(N).ToList();
                List<double?> dev = series.StandardDeviation(N).ToList();

                var smaRows = smaData.Rows;
                var stdRows = stdData.Rows;
                var lowerBBRows = lowerBBData.Rows;
                var upperBBRows = upperBBData.Rows;

                for (int i = 0; i < eodData.Rows.Count; i++)
                {
                    smaRows[i][colname] = sma[i];
                    stdRows[i][colname] = dev[i];
                    lowerBBRows[i][colname] = sma[i] - K * dev[i];
                    upperBBRows[i][colname] = sma[i] + K * dev[i];
                }
            }
        }


        public DataTable runIntraday(string rundate)
        {
            // local variables
            orderPosition   = new Dictionary<string, int>();
            orderProfit     = new Dictionary<string, double>();
            orderEntryPrice = new Dictionary<string, double>();

            intraData = new DataTable();
            intraData = Utility.ReadExcel(@"..\..\" + rundate + ".csv");
            intraData.Columns["Column1"].ColumnName = "Time";
            orders = intraData.Copy();

            // initialize the orders table
            var ordercols = orders.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            var tickers = ordercols.Where((value, index) => index >= 1 && index <= ordercols.Count - 1).ToList();
            foreach (DataRow orderrow in orders.Rows)
            {
                foreach (string ticker in tickers)
                {
                    orderrow[ticker] = null;
                }
            }

            // initialize order position and profit tracking
            foreach (string ticker in tickers)
            {
                orderPosition.Add(ticker, 0);
                orderProfit.Add(ticker, 0);
                orderEntryPrice.Add(ticker, -1);
            }

            // get the target levels for the runday based on previous day's Bollinger Bands
            var dateseries = eodData.AsEnumerable().Select(r => r.Field<string>("Date")).ToList();
            int indexeod = dateseries.IndexOf(rundate.Replace(".", "-"));
            rundaySMA = new Dictionary<string, double>();
            rundayLBB = new Dictionary<string, double>();
            rundayUBB = new Dictionary<string, double>();
            foreach (string ticker in tickers)
            {
                rundaySMA.Add(ticker, (double)smaData.Rows[indexeod - 1][ticker + " US"]);
                rundayLBB.Add(ticker, (double)lowerBBData.Rows[indexeod - 1][ticker + " US"]);
                rundayUBB.Add(ticker, (double)upperBBData.Rows[indexeod - 1][ticker + " US"]);
            }

            // REPLAY the intraday data and fire the orders as needed
            // Square off at end of day
            for (int indexintra = 0; indexintra < intraData.Rows.Count - 1; indexintra++)
            {
                DataRow dr = intraData.Rows[indexintra];
                foreach (string ticker in tickers)
                {
                    if (orderPosition[ticker] == 0)
                    {
                        if ((double)dr[ticker] <= rundayLBB[ticker])
                        {
                            orderPosition[ticker] = 1;
                            orders.Rows[indexintra][ticker] = "BUY";
                            orderEntryPrice[ticker] = (double)dr[ticker];
                        }
                        else if ((double)dr[ticker] >= rundayUBB[ticker])
                        {
                            orderPosition[ticker] = -1;
                            orders.Rows[indexintra][ticker] = "SELL";
                            orderEntryPrice[ticker] = (double)dr[ticker];
                        }
                    }
                    else if (orderPosition[ticker] == 1)
                    {
                        if ((double)dr[ticker] >= rundaySMA[ticker])
                        {
                            orderPosition[ticker] = 0;
                            orders.Rows[indexintra][ticker] = "SELL";
                            orderProfit[ticker] += (double)dr[ticker] - orderEntryPrice[ticker];
                        }
                    }
                    else if (orderPosition[ticker] == -1)
                    {
                        if ((double)dr[ticker] <= rundaySMA[ticker])
                        {
                            orderPosition[ticker] = 0;
                            orders.Rows[indexintra][ticker] = "BUY";
                            orderProfit[ticker] += -1 * ((double)dr[ticker] - orderEntryPrice[ticker]);
                        }
                    }
                }
            }

            // Square off positions at End of Day
            DataRow drlast = intraData.Rows[intraData.Rows.Count - 1];
            foreach (string ticker in tickers)
            {
                if (orderPosition[ticker] == 1)
                {
                    orderPosition[ticker] = 0;
                    orders.Rows[intraData.Rows.Count - 1][ticker] = "SELL";
                    orderProfit[ticker] += (double)drlast[ticker] - orderEntryPrice[ticker];
                }
                else if (orderPosition[ticker] == -1)
                {
                    orderPosition[ticker] = 0;
                    orders.Rows[intraData.Rows.Count - 1][ticker] = "BUY";
                    orderProfit[ticker] += -1 * ((double)drlast[ticker] - orderEntryPrice[ticker]);
                }
            }
            return orders;
        }

        public string Checkordertime(DataTable ord, string orderdate, string ticker, string buysell)
        {
            var foundRows = ord.AsEnumerable().Where(dr => dr.Field<string>(ticker) == buysell);
            foreach (var row in foundRows)
            {
                return (string)row[0];
            }
            return null;
        }



        public static void Main(string[] args)
        {
            var runObject = new Bollinger("last.csv");

            List<string> rundates = new List<string> { "2022.05.02", "2022.05.03", "2022.05.04", "2022.05.05", "2022.05.06" };
            foreach (string rundate in rundates)
            {
                var runorders = runObject.runIntraday(rundate);
                Console.WriteLine("Orders for date " + rundate);
                Utility.printDataTable(runorders);
                Console.WriteLine("\n------------x-------------\n");
                Utility.csvDataTable(runorders, @"..\..\" + "orders_" + rundate + ".csv");
            }       
        }
    }
}
