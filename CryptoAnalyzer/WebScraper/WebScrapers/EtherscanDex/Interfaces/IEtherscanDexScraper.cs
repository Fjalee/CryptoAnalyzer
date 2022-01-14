﻿using System.Collections.Generic;
using System.Threading.Tasks;

namespace WebScraper.WebScrapers
{
    public interface IDexScraper
    {
        Task<List<DexRow>> ScrapeCurrentPageTable();
        void GoToNextPage();
        void Initialize(string url);
    }
}