using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using NLog;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeIntegration.Service
{
    static class Program
    {
        static Logger log = LogManager.GetCurrentClassLogger();
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "-debug")
            {
                try
                {
                    Service1 srv = new Service1();
                    srv.Start(args);
                    Console.WriteLine("Hit enter to exit");
                    Console.ReadLine();
                    srv.Stop();
                }
                catch (Exception ex)
                {
                    log.Error("Error: {0}\n.Enter to exit.", ex);
                    Console.ReadLine();
                }
                return;
            }
            else
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] 
			    { 
				    new Service1() 
			    };
                ServiceBase.Run(ServicesToRun);
            }
        }

        static void Test()
        {
            
        }
    }
}
