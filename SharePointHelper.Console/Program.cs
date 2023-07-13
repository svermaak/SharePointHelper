using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointHelper.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            Core.Config.ProcessRules();
            //Core.PDFThumbnail.ProcessList("http://sp2013dev", "http://sp2013dev","/", "PDFLibrary - Entries", "PDFLibrary - Thumbnails");
            //http://sp2013dev|http://sp2013dev|/|TestDocs|Status|Approved
            //System.Console.ReadKey();

        }
    }
}
