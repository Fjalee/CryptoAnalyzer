﻿using CsvHelper;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace CryptoAnalyzer
{
    public class CsvOutput : IOutput
    {
        public void CreateFile(string pathName, string fileName, List<WebScraper.Transaction> list)
        {
            var fullPath = pathName + '/' + fileName + ".csv";


            Directory.CreateDirectory(pathName);
            using (var writer = File.AppendText(fullPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(list);
            }
        }
    }
}