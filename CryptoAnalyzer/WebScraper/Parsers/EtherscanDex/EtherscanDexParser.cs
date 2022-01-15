﻿using WebScraper.WebScrapers;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Diagnostics;
using System;
using WebScraper.WebScrapers.EtherscanDex;
using System.Threading.Tasks;

namespace WebScraper.Parsers
{
    public class EtherscanDexParser : IDexTableParser
    {
        private readonly IParserCommon _parserCommon;
        private readonly IEtherscanApiServices _etherscanApiServices;

        public EtherscanDexParser(IParserCommon parserCommon, IEtherscanApiServices etherscanApiServices)
        {
            _parserCommon = parserCommon;
            _etherscanApiServices = etherscanApiServices;
        }

        public IElement GetTable(IHtmlDocument page)
        {
            IElement current = null;
            try
            {
                current = page.Body.Children[10];
                _parserCommon.StepIfMatches(ref current, "doneloadingframe", current.Id, current.Children[2]);
                _parserCommon.StepIfMatches(ref current, "table-responsive mb-2 mb-md-0", current.ClassName, current.Children[0]);
                _parserCommon.StepIfMatches(ref current, "table table-hover", current.ClassName, current.Children[1]);
                _parserCommon.StepIfMatches(ref current, "tbody", current.LocalName, current);
            }
            catch
            {
                State.ExitAndLog(new StackTrace());
            }

            return current;
        }

        public async Task<DexRow> ParseRow(IElement rowHtml)
        {
            var row = new DexRow();

            try
            {
                if (10 != rowHtml.Children.Length)
                {
                    State.ExitAndLog(new StackTrace());
                }

                row.TxnHash = ParseTransactionHash(rowHtml);
                row.TxnDate = ParseDate(rowHtml);
                row.Action = ParseAction(rowHtml);

                var txnDetails = await _etherscanApiServices.GetTransactionDetailsAsync(row.TxnHash);

                row.SellerHash = txnDetails.From;
                row.BuyerHash = txnDetails.To;
            }
            catch(Exception e)
            {
                State.ExitAndLog(new StackTrace());
            }

            return row;
        }

        private DateTime ParseDate(IElement rowHtml)
        {
            var dateRow = rowHtml.Children[2];
            var parsedTxnDate = _parserCommon.GetDataIfMatches(dateRow.Children[0].InnerHtml, "showDate ", dateRow.ClassName);
            return DateTime.Parse(parsedTxnDate);
        }

        private DexAction ParseAction(IElement rowHtml)
        {
            var isBuyRow = rowHtml.Children[4];
            var parsedAction = isBuyRow.Children.Length == 0 ?
                _parserCommon.GetDataIfMatches(isBuyRow.InnerHtml, "td", isBuyRow.LocalName) :
                _parserCommon.GetDataIfMatches(isBuyRow.Children[0].InnerHtml, "span", isBuyRow.Children[0].LocalName);


            var successfulParse = Enum.TryParse(parsedAction, out DexAction result);
            if (!successfulParse)
            {
                throw new InvalidOperationException($"row Action was expected to be equal to Buy or Sell");
            }
            
            return result;
        }
        private string ParseTransactionHash(IElement rowHtml)
        {
            var dateRow = rowHtml.Children[1].Children[0];
            return _parserCommon.GetDataIfMatches(dateRow.TextContent, "hash-tag text-truncate myFnExpandBox_searchVal", dateRow.ClassName);
        }
    }
}
