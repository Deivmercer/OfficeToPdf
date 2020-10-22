using System;
using AntiVirus;

namespace FileConverter
{
    public class TrovaVirus
    {

        /// <summary>
        ///		Inizializzazione del logger
        /// </summary>
        public static log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public ScanResult ControllaVirus(String pathDaControllare)
        {
            //Istanzio lo scanner antivirus
            var scanner = new Scanner();

            //Scannerizzo il file e lo ripulisco dai virus
            ScanResult result = scanner.ScanAndClean(pathDaControllare);

            return result;
        }
    }
}
