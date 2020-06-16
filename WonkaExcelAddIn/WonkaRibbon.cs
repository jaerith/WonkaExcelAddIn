using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

using Wonka.BizRulesEngine;
using Wonka.BizRulesEngine.RuleTree;
using Wonka.Eth.Init;
using Wonka.MetaData;
using Wonka.Product;

namespace WonkaExcelAddIn
{
    public partial class WonkaRibbon
    {
		#region CONSTANTS

		public const string CONST_ACCT_PUBLIC_KEY   = "0x12890D2cce102216644c59daE5baed380d84830c";
		public const string CONST_ACCT_PASSWORD     = "0xb5b1870957d373ef0eeffecc6e4812c0fd08f554b37b233526acc331bf1544f7";

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

		private string               currRulesUrl   = String.Format("{0}/{1}", CONST_INFURA_IPFS_GATEWAY_URL, CONST_RULES_FILE_IPFS_KEY);
		private string               wonkaRules     = "";
		private IMetadataRetrievable metadataSource = new Wonka.BizRulesEngine.Samples.WonkaBreMetadataTestSource();
		private WonkaBizRulesEngine  rulesEngine    = null;

		private WonkaRefEnvironment refEnvHandle = null;

		private void WonkaRibbon_Load(object sender, RibbonUIEventArgs e)
        {
			refEnvHandle =
				WonkaRefEnvironment.CreateInstance(false, metadataSource);

			WonkaRefAttr AccountIDAttr       = refEnvHandle.GetAttributeByAttrName("BankAccountID");
			WonkaRefAttr AccountNameAttr     = refEnvHandle.GetAttributeByAttrName("BankAccountName");
			WonkaRefAttr AccountStsAttr      = refEnvHandle.GetAttributeByAttrName("AccountStatus");
			WonkaRefAttr AccountCurrValAttr  = refEnvHandle.GetAttributeByAttrName("AccountCurrValue");
			WonkaRefAttr AccountTypeAttr     = refEnvHandle.GetAttributeByAttrName("AccountType");
			WonkaRefAttr AccountCurrencyAttr = refEnvHandle.GetAttributeByAttrName("AccountCurrency");
			WonkaRefAttr RvwFlagAttr         = refEnvHandle.GetAttributeByAttrName("AuditReviewFlag");
			WonkaRefAttr CreationDtAttr      = refEnvHandle.GetAttributeByAttrName("CreationDt");

			using (var client = new System.Net.Http.HttpClient())
			{
				wonkaRules = client.GetStringAsync(currRulesUrl).Result;
			}

			rulesEngine = 
				new WonkaBizRulesEngine(new StringBuilder(wonkaRules));
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

		private WonkaEthRulesEngine AssembleWonkaEthEngine(string psWonkaRules)
		{
			refEnvHandle =
				WonkaRefEnvironment.CreateInstance(false, metadataSource);

			string sContractAddr  = "";
			string sContractABI   = "";
			string sGetMethodName = "";
			string sSetMethodName = "";

			WonkaBizSource.RetrieveDataMethod retrieveMethod = null;

			var SourceMap = new Dictionary<string, WonkaBizSource>();
			foreach (var TmpAttr in refEnvHandle.AttrCache)
			{
				var TmpSource =
					new WonkaBizSource(TmpAttr.AttrName,
									   CONST_ACCT_PUBLIC_KEY,
									   CONST_ACCT_PASSWORD,
									   sContractAddr,
									   sContractABI,
									   sGetMethodName,
									   sSetMethodName,
									   retrieveMethod);

				SourceMap[TmpAttr.AttrName] = TmpSource;
			}

			WonkaEthEngineInitialization EngineInit =
				new WonkaEthEngineInitialization() { EthSenderAddress = CONST_ACCT_PUBLIC_KEY, 
					                                 EthPassword = CONST_ACCT_PASSWORD, 
					                                 Web3HttpUrl = CONST_TEST_INFURA_URL };

			return new WonkaEthRulesEngine(new StringBuilder(psWonkaRules), SourceMap, EngineInit, metadataSource, false);
		}

		private Dictionary<string,string> DisassembleProduct(Wonka.Product.WonkaProduct poProduct)
		{
			var DataSnapshot = new Dictionary<string, string>();

			refEnvHandle.AttrCache.ForEach(x => DataSnapshot[x.AttrName] = poProduct.GetAttributeValue(x));

			return DataSnapshot;
		}

		public string GetErrors(Wonka.BizRulesEngine.Reporting.WonkaBizRuleTreeReport report)
		{
			var ErrorReport = new StringBuilder();

			foreach (var ReportNode in report.GetRuleSetSevereFailures())
			{
				if (ReportNode.RuleResults.Count > 0)
				{
					foreach (var RuleReportNode in ReportNode.RuleResults)
					{
						if (ErrorReport.Length > 0)
							ErrorReport.Append("\n");

						ErrorReport.Append(RuleReportNode.VerboseError.Replace("/", ""));
					}
				}
				else
					ErrorReport.Append(ReportNode.ErrorDescription);
			}

			return ErrorReport.ToString();
		}

		private void Validate_Click(object sender, RibbonControlEventArgs e)
        {
			try
			{
				var thisCell     = Globals.ThisAddIn.GetActiveCell();
				var currAttrData = Globals.ThisAddIn.GetCurrentAttributeData();

				WonkaProduct currProduct = AssembleProduct(currAttrData);

				var report = rulesEngine.Validate(currProduct);

				if (report.GetRuleSetSevereFailureCount() == 0)
				{
					var postAttrData = DisassembleProduct(currProduct);

					Globals.ThisAddIn.SetCurrentAttributeData(postAttrData);

					MessageBox.Show("SUCCESS!");
				}
				else
				{
					string sErrors = GetErrors(report);
					MessageBox.Show("ERROR!  " + sErrors);
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
