using DMobiAnalysis.Models;
using Newtonsoft.Json;
using Novacode;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;

namespace DMobiAnalysis
{
    class Program
    {
        static void Main(string[] args)
        {
            DocumentProcessor documentProcessor = new DocumentProcessor();
            documentProcessor.ProcessDocuments(@"C:\Users\carel\Desktop\New folder\GS01");
        }
    }
}
