using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace ServicioPLCSQL
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {
//#if DEBUG
//            ServicioPLCSQL myservice = new ServicioPLCSQL();
//            myservice.OnDebug();
//            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
//#else

//#endif

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new ServicioPLCSQL() 
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
