﻿using System.Collections.Generic;
using System.Threading.Tasks;

namespace WebScraper.WebScrapers
{
    public interface IDexScraper
    {
        Task<List<DexRow>> ScrapeTable(string url);
    }
}