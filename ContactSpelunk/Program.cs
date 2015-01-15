
namespace ContactSpelunk
{
    using Microsoft.Office.Interop.Outlook;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    class Program
    {
        static void Main(string[] args)
        {
            Application app = null;
            _NameSpace ns = null;

            Dictionary<string, int> totalCounts = new Dictionary<string, int>();
            float totalProcessed = 0;
            float filteredCount = 0;
            Stopwatch sw = new Stopwatch();
            sw.Start();

            try
            {
                app = new Application();
                ns = app.GetNamespace("MAPI");
                ns.Logon(null, null, false, false);

                AddressList GAL = ns.AddressLists["Global Address List"];

                Console.WriteLine("Total count: {0}", GAL.AddressEntries.Count);

                foreach (AddressEntry oEntry in GAL.AddressEntries)
                //Parallel.ForEach<AddressEntry>(GAL.AddressEntries.Cast<AddressEntry>(), oEntry =>
                {
                    totalProcessed++;
                    if (totalProcessed % 100 == 0)
                    {
                        Console.Write("\rProcssing ... {0} of {1} ({2}%). Processing @ {3} ps, Expected completion: [{4}]", 
                            totalProcessed,
                            GAL.AddressEntries.Count,
                            ((totalProcessed * 100.00) / GAL.AddressEntries.Count).ToString("000.0000"),
                            (totalProcessed/sw.Elapsed.TotalSeconds).ToString("00.00"),
                            DateTime.Now.AddSeconds(sw.Elapsed.TotalSeconds * (GAL.AddressEntries.Count-totalProcessed)/totalProcessed));
                    }

                    if (oEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry)
                    {
                        //Task.Factory.StartNew(() =>
                        {
                            ExchangeUser contact = oEntry.GetExchangeUser();

                            if (contact != null)
                            {

                                filteredCount++;
                                if (!String.IsNullOrEmpty(contact.JobTitle))
                                {
                                    if (!totalCounts.ContainsKey(contact.JobTitle))
                                    {
                                        totalCounts.Add(contact.JobTitle, 0);
                                    }

                                    totalCounts[contact.JobTitle]++;
                                }
                            }
                        }
                        //);
                    }
                }
                //);

                sw.Stop();
                Console.WriteLine("Processed {0} entires (Filtered Count: {1}) in {2}",
                    totalProcessed, 
                    filteredCount,
                    sw.Elapsed.ToString());
                
                foreach (KeyValuePair<string, int> kvp in totalCounts)
                {
                    Console.WriteLine("{0}\t{1}", kvp.Value, kvp.Key);
                }                
            }
            catch(System.Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
