using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GuetiTech;


namespace PdfConvertorTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Out.WriteLine("Start Program");

            string fileToConvert = "C:\\Temp\\Test.rtf";

            if (args.Length == 1)
            {
                fileToConvert = args[0];
            }

            try
            {

                Console.Out.WriteLine("File to convert: " + fileToConvert);

                byte[] fileToConvertBytes = System.IO.File.ReadAllBytes(fileToConvert);

                Console.Out.WriteLine("Rtf Bytes: " + Convert.ToBase64String(fileToConvertBytes));

                string base64PdfContent = RtfToPdfConverter.ConvertRtfToPdf("Unique", fileToConvertBytes, false);

                Console.Out.WriteLine("Base64 Bytes: " + base64PdfContent);

                System.IO.File.WriteAllBytes(fileToConvert + ".pdf", Convert.FromBase64String(base64PdfContent));

                Console.Out.WriteLine("End Program");
            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e.Message);
            }
        }
    }
}
