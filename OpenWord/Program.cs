using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace OpenWord
{
    class Program
    {
        static void Main(string[] args)
        {
            // use reg file to add custom browser protocols 
            // open file from webdav server on iis
            //msword://https://localhost/webdav/Test1.docx

            File.WriteAllLines(@"c:\temp\openword.txt", args);
            ProcessStartInfo start = new ProcessStartInfo();
            start.Arguments = args[0].Substring(9, args[0].Length - 9);
            start.FileName = "winword.exe";
            start.WindowStyle = ProcessWindowStyle.Maximized;
            start.CreateNoWindow = true;
            Process proc = Process.Start(start);
        }
    }
}
