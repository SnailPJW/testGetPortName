using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleComPort
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //{4d36e978-e325-11ce-bfc1-08002be10318}為設備類別port（端口（COM&LPT））的GUID
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2",
                    "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\"");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    string name = queryObj.GetPropertyValue("Name").ToString();
                    string desc = queryObj.GetPropertyValue("Description").ToString();
                    Console.WriteLine("{0} : {1}", name, desc);
                }
            }
            catch (Exception ex)
            {
                //Log.w
            }
            finally
            {
                Console.ReadLine();
            }

        }
    }
}
