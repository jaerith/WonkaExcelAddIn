using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

using Wonka.BizRulesEngine;
using Wonka.MetaData;

namespace WonkaExcelAddIn
{
    public partial class WonkaRibbon
    {
		#region CONSTANTS

		public const string CONST_INFURA_IPFS_GATEWAY_URL     = "https://ipfs.infura.io/ipfs/";
		public const string CONST_INFURA_IPFS_API_GATEWAY_URL = "https://ipfs.infura.io:5001/7238211010344719ad14a89db874158c/api/";
		public const string CONST_TEST_INFURA_KEY             = "7238211010344719ad14a89db874158c";
		public const string CONST_TEST_INFURA_URL             = "https://mainnet.infura.io/v3/7238211010344719ad14a89db874158c";
		public const string CONST_ETH_FNDTN_EOA_ADDRESS       = "0xde0b295669a9fd93d5f28d9ec85e40f4cb697bae";
		public const string CONST_DAI_TOKEN_CTRCT_ADDRESS     = "0x89d24a6b4ccb1b6faa2625fe562bdd9a23260359";
		public const string CONST_MAKER_ERC20_CTRCT_ADDRESS   = "0x9f8f72aa9304c8b593d555f12ef6589cc3a579a2";
		public const string CONST_RULES_FILE_IPFS_KEY         = "QmXcsGDQthxbGW8C3Sx9r4tV9PGSj4MxJmtXF7dnXN5XUT";
		public const string CONST_VAT_RULES_FILE_IPFS_KEY     = "QmPrZ9959c7SzzqdLkVgX28xM7ZrqLeT3ydvRAHCaL1Hsn";
		public const string CONST_METADATA_FILE_IPFS_KEY      = "QmYLc2Ej17hHBwz8zmjm4a42h4fbm68hzzpEmQKfVKgYrU";
		public const string CONST_VAT_METADATA_FILE_IPFS_KEY  = "QmagCzTxsrbPWwze3pDhVpYeWYB9E2LFk1FNqUngFyZqSN";

		#endregion

		private string              currRulesUrl = String.Format("{0}/{1}", CONST_INFURA_IPFS_GATEWAY_URL, CONST_RULES_FILE_IPFS_KEY);
		private WonkaBizRulesEngine rulesEngine  = null;

		private void WonkaRibbon_Load(object sender, RibbonUIEventArgs e)
        {
			WonkaRefEnvironment WonkaRefEnv =
				WonkaRefEnvironment.CreateInstance(false, new Wonka.BizRulesEngine.Samples.WonkaBreMetadataTestSource());

			WonkaRefAttr AccountIDAttr       = WonkaRefEnv.GetAttributeByAttrName("BankAccountID");
			WonkaRefAttr AccountNameAttr     = WonkaRefEnv.GetAttributeByAttrName("BankAccountName");
			WonkaRefAttr AccountStsAttr      = WonkaRefEnv.GetAttributeByAttrName("AccountStatus");
			WonkaRefAttr AccountCurrValAttr  = WonkaRefEnv.GetAttributeByAttrName("AccountCurrValue");
			WonkaRefAttr AccountTypeAttr     = WonkaRefEnv.GetAttributeByAttrName("AccountType");
			WonkaRefAttr AccountCurrencyAttr = WonkaRefEnv.GetAttributeByAttrName("AccountCurrency");
			WonkaRefAttr RvwFlagAttr         = WonkaRefEnv.GetAttributeByAttrName("AuditReviewFlag");
			WonkaRefAttr CreationDtAttr      = WonkaRefEnv.GetAttributeByAttrName("CreationDt");

			string sWonkaRules = "";

			using (var client = new System.Net.Http.HttpClient())
			{
				sWonkaRules = client.GetStringAsync(currRulesUrl).Result;
			}

			rulesEngine = new WonkaBizRulesEngine(new StringBuilder(sWonkaRules));
        }

        private void Validate_Click(object sender, RibbonControlEventArgs e)
        {
            var thisCell = Globals.ThisAddIn.GetActiveCell();

			var CurrAttrData = Globals.ThisAddIn.GetCurrentAttributeData();

			// MessageBox.Show("This cell’s address is: " + thisCell.Address + " — And it’s value is: " + thisCell.Value);
		}
    }
}
