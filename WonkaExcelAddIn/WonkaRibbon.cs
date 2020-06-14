using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

using Wonka.BizRulesEngine;
using Wonka.MetaData;
using Wonka.Product;

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

		private WonkaRefEnvironment refEnvHandle = null;

		private void WonkaRibbon_Load(object sender, RibbonUIEventArgs e)
        {
			refEnvHandle =
				WonkaRefEnvironment.CreateInstance(false, new Wonka.BizRulesEngine.Samples.WonkaBreMetadataTestSource());

			WonkaRefAttr AccountIDAttr       = refEnvHandle.GetAttributeByAttrName("BankAccountID");
			WonkaRefAttr AccountNameAttr     = refEnvHandle.GetAttributeByAttrName("BankAccountName");
			WonkaRefAttr AccountStsAttr      = refEnvHandle.GetAttributeByAttrName("AccountStatus");
			WonkaRefAttr AccountCurrValAttr  = refEnvHandle.GetAttributeByAttrName("AccountCurrValue");
			WonkaRefAttr AccountTypeAttr     = refEnvHandle.GetAttributeByAttrName("AccountType");
			WonkaRefAttr AccountCurrencyAttr = refEnvHandle.GetAttributeByAttrName("AccountCurrency");
			WonkaRefAttr RvwFlagAttr         = refEnvHandle.GetAttributeByAttrName("AuditReviewFlag");
			WonkaRefAttr CreationDtAttr      = refEnvHandle.GetAttributeByAttrName("CreationDt");

			string sWonkaRules = "";

			using (var client = new System.Net.Http.HttpClient())
			{
				sWonkaRules = client.GetStringAsync(currRulesUrl).Result;
			}

			rulesEngine = new WonkaBizRulesEngine(new StringBuilder(sWonkaRules));

			/**
			 ** NOTE: Now set the data on the worksheet
			 **
			var sampleData = new Dictionary<string, string>();
			sampleData[AccountIDAttr.AttrName]   = "123456789";
			sampleData[AccountNameAttr.AttrName] = "JohnSmithFirstCheckingAccount";
			Globals.ThisAddIn.SetCurrentAttributeData(sampleData);
	         **/
		}

		private WonkaProduct AssembleProduct(Dictionary<string, string> poAttrData)
		{
			var NewProduct = new Wonka.Product.WonkaProduct();

			foreach (string sTmpAttrName in poAttrData.Keys)
			{
				WonkaRefAttr TargetAttr = refEnvHandle.GetAttributeByAttrName(sTmpAttrName);

				NewProduct.SetAttribute(TargetAttr, poAttrData[sTmpAttrName]);
			}

			return NewProduct;
		}

		private void Validate_Click(object sender, RibbonControlEventArgs e)
        {
			try
			{
				var thisCell = Globals.ThisAddIn.GetActiveCell();

				var currAttrData = Globals.ThisAddIn.GetCurrentAttributeData();

				WonkaProduct currProduct = AssembleProduct(currAttrData);

				var report = rulesEngine.Validate(currProduct);

				if (report.GetRuleSetSevereFailureCount() == 0)
					MessageBox.Show("SUCCESS!");
				else
				{
					MessageBox.Show("ERROR!  [" + report.GetRuleSetSevereFailureCount() + "] severe rulesets failed.");
				}
			}
			catch (WonkaBizRuleException bizEx)
			{
				MessageBox.Show("ERROR!  Wonka Exception: " + bizEx);
			}
			catch (Exception ex)
			{
				MessageBox.Show("ERROR!  Exception: " + ex);
			}
		}
	}
}
