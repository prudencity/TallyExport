using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Configuration;

namespace TallyExport
{
    public partial class form_TallyExport : Form
    {

        string value = ConfigurationManager.AppSettings["TallyExportDir"];

        public form_TallyExport()
        {
            InitializeComponent();
        }

        class_DataAccess ObjData = new class_DataAccess();                      //Object to access the Data from the database.

        private void btn_Export_Click(object sender, EventArgs e)
        {
            DateTime userEnteredDate = DatePicker.Value;

            bool checkFile1 = false;
            bool checkFile2 = false;

            String input_date = userEnteredDate.ToString("yyyyMMdd");
            
            Export_Ledgers(input_date);
            Export_Transactions(input_date);

            checkFile1 = System.IO.File.Exists(String.IsNullOrEmpty(value) ? "C:\\Tally" : value + "\\m-" + input_date);
            checkFile2 = System.IO.File.Exists(String.IsNullOrEmpty(value) ? "C:\\Tally" : value + "\\v-" + input_date);

            if (checkFile1 == true && checkFile2 == true)
            {
                MessageBox.Show("Files(s) Exported Successfully !", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            else if (checkFile1 == false)
            {
                MessageBox.Show("Ledgers not exported ! Please try again !", "Ledgers Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
            else if (checkFile2 == false)
            {
                MessageBox.Show("Vouchers not exported ! Please try again !", "Vouchers Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
            else
            {
                MessageBox.Show("Files not exported ! Please try again !", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }

        }

        #region Export all Ledgers in XML Format for Tally
        private void Export_Ledgers(String input_date)
        {
            XmlDocument doc = new XmlDocument();
            XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.AppendChild(docNode);

            XmlNode envelopeNode = doc.CreateElement("ENVELOPE");   //ENVELOPE
            doc.AppendChild(envelopeNode);

            XmlNode headerNode = doc.CreateElement("HEADER");       //HEADER
            envelopeNode.AppendChild(headerNode);

            XmlNode tallyRequestNode = doc.CreateElement("TALLYREQUEST");   //TALLY REQUEST
            tallyRequestNode.AppendChild(doc.CreateTextNode("Import Data"));
            headerNode.AppendChild(tallyRequestNode);


            XmlNode bodyNode = doc.CreateElement("BODY");           //BODY
            envelopeNode.AppendChild(bodyNode);

            XmlNode importDataNode = doc.CreateElement("IMPORTDATA");       //IMPORT DATA
            bodyNode.AppendChild(importDataNode);

            XmlNode requestDescNode = doc.CreateElement("REQUESTDESC");     //REQUEST DESC
            importDataNode.AppendChild(requestDescNode);

            XmlNode reportNameNode = doc.CreateElement("REPORTNAME");       //REPORT NAME
            reportNameNode.AppendChild(doc.CreateTextNode("All Masters"));              //Name of the Report
            requestDescNode.AppendChild(reportNameNode);

            XmlNode staticVariablesNode = doc.CreateElement("STATICVARIABLES");
            requestDescNode.AppendChild(staticVariablesNode);

            XmlNode svCurrentCompnayNode = doc.CreateElement("SVCURRENTCOMPANY");       //SVCURRENTCOMPANY

            svCurrentCompnayNode.AppendChild(doc.CreateTextNode("company_name"));      //Datapull 
            staticVariablesNode.AppendChild(svCurrentCompnayNode);

            XmlNode requestDataNode = doc.CreateElement("REQUESTDATA");
            importDataNode.AppendChild(requestDataNode);

            XmlNode tallyMessageNode = doc.CreateElement("TALLYMESSAGE");
            XmlAttribute tallyMessageAttribute = doc.CreateAttribute("xmlns:UDF");
            tallyMessageAttribute.Value = "TallyUDF";
            tallyMessageNode.Attributes.Append(tallyMessageAttribute);
            requestDataNode.AppendChild(tallyMessageNode);

            XmlNode currencyNode = doc.CreateElement("CURRENCY");
            XmlAttribute currencyAttribute1 = doc.CreateAttribute("NAME");
            currencyAttribute1.Value = "Rs.";
            currencyNode.AppendChild(doc.CreateTextNode(""));
            XmlAttribute currencyAttribute2 = doc.CreateAttribute("RESERVEDNAME");
            currencyAttribute2.Value = "";
            currencyNode.Attributes.Append(currencyAttribute1);
            currencyNode.Attributes.Append(currencyAttribute2);
            tallyMessageNode.AppendChild(currencyNode);

            XmlNode mailingNameNode = doc.CreateElement("MAILINGNAME");
            mailingNameNode.AppendChild(doc.CreateTextNode("Indian Rupees"));
            currencyNode.AppendChild(mailingNameNode);

            XmlNode expandedSymbolNode = doc.CreateElement("EXPANDEDSYMBOL");
            expandedSymbolNode.AppendChild(doc.CreateTextNode("Indian Rupees"));
            currencyNode.AppendChild(expandedSymbolNode);

            XmlNode decimalSymbolNode = doc.CreateElement("DECIMALSYMBOL");
            decimalSymbolNode.AppendChild(doc.CreateTextNode("paise"));
            currencyNode.AppendChild(decimalSymbolNode);

            XmlNode originalSymbolNode = doc.CreateElement("ORIGINALSYMBOL");
            originalSymbolNode.AppendChild(doc.CreateTextNode("Rs."));
            currencyNode.AppendChild(originalSymbolNode);

            XmlNode isSuffixNode = doc.CreateElement("ISSUFFIX");
            isSuffixNode.AppendChild(doc.CreateTextNode("No"));
            currencyNode.AppendChild(isSuffixNode);

            XmlNode hasSpaceNode = doc.CreateElement("HASSPACE");
            hasSpaceNode.AppendChild(doc.CreateTextNode("Yes"));
            currencyNode.AppendChild(hasSpaceNode);

            XmlNode inMillionsNode = doc.CreateElement("INMILLIONS");
            inMillionsNode.AppendChild(doc.CreateTextNode("No"));
            currencyNode.AppendChild(inMillionsNode);

            XmlNode decimalPlacesNode = doc.CreateElement("DECIMALPLACES");
            decimalPlacesNode.AppendChild(doc.CreateTextNode(" 2"));
            currencyNode.AppendChild(decimalPlacesNode);

            XmlNode decimalPlacesForPrintingNode = doc.CreateElement("DECIMALPLACESFORPRINTING");
            decimalPlacesForPrintingNode.AppendChild(doc.CreateTextNode(" 2"));
            currencyNode.AppendChild(decimalPlacesForPrintingNode);

            String S1 = "select P.P_Name, P.P_Group, AG.ag_name from PartyMaster as P, AccountGroups as AG where AG.ag_code = P.P_Group AND P.P_Name IS NOT NULL AND P.P_Name <> '';";
            DataTable Dt1 = ObjData.GetDataTable(S1);

            for (int i = 0; i < Dt1.Rows.Count; i++)
            {

                String T1 = Dt1.Rows[i]["P_Name"].ToString();
                String T2 = Dt1.Rows[i]["P_Group"].ToString();
                String T3 = Dt1.Rows[i]["ag_name"].ToString();

                XmlNode tallyMessage1Node = doc.CreateElement("TALLYMESSAGE");
                XmlAttribute tallyMessage1Attribute = doc.CreateAttribute("xmlns:UDF");
                tallyMessage1Attribute.Value = "TallyUDF";
                tallyMessage1Node.Attributes.Append(tallyMessage1Attribute);
                requestDataNode.AppendChild(tallyMessage1Node);

                XmlNode ledgerNode = doc.CreateElement("LEDGER");
                XmlAttribute ledgerAttribute = doc.CreateAttribute("NAME");
                ledgerAttribute.Value = T1;
                XmlAttribute ledgerAttribute1 = doc.CreateAttribute("RESERVEDNAME");
                ledgerAttribute1.Value = "";
                ledgerNode.Attributes.Append(ledgerAttribute);
                ledgerNode.Attributes.Append(ledgerAttribute1);
                ledgerNode.AppendChild(doc.CreateTextNode(""));

                tallyMessage1Node.AppendChild(ledgerNode);


                XmlNode mailingNameListNode = doc.CreateElement("MAILINGNAME.LIST");
                XmlAttribute mailingNameListAttribute = doc.CreateAttribute("TYPE");
                mailingNameListAttribute.Value = "String";
                mailingNameListNode.AppendChild(doc.CreateTextNode(""));
                mailingNameListNode.Attributes.Append(mailingNameListAttribute);
                ledgerNode.AppendChild(mailingNameListNode);

                XmlNode mailingListMailingNameNode = doc.CreateElement("MAILINGNAME");
                mailingListMailingNameNode.AppendChild(doc.CreateTextNode(T1));
                mailingNameListNode.Attributes.Append(mailingNameListAttribute);
                mailingNameListNode.AppendChild(mailingListMailingNameNode);

                XmlNode currencyNameNode = doc.CreateElement("CURRENCYNAME");
                currencyNameNode.AppendChild(doc.CreateTextNode("Rs."));
                ledgerNode.AppendChild(currencyNameNode);

                XmlNode parentNode = doc.CreateElement("PARENT");
                parentNode.AppendChild(doc.CreateTextNode(T3));
                ledgerNode.AppendChild(parentNode);

                XmlNode exciseLedgerClassificationNode = doc.CreateElement("EXCISELEDGERCLASSIFICATION");
                exciseLedgerClassificationNode.AppendChild(doc.CreateTextNode("Default"));
                ledgerNode.AppendChild(exciseLedgerClassificationNode);

                XmlNode isBillwiseOnNode = doc.CreateElement("ISBILLWISEON");
                isBillwiseOnNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isBillwiseOnNode);

                XmlNode isCostCentresOnNode = doc.CreateElement("ISCOSTCENTRESON");
                isCostCentresOnNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isCostCentresOnNode);

                XmlNode isInterestOnNode = doc.CreateElement("ISINTERESTON");
                isInterestOnNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isInterestOnNode);

                XmlNode allowInMobileNode = doc.CreateElement("ALLOWINMOBILE");
                allowInMobileNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(allowInMobileNode);

                XmlNode isCondensedNode = doc.CreateElement("ISCONDENSED");
                isCondensedNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isCondensedNode);

                XmlNode affectsStockNode = doc.CreateElement("AFFECTSSTOCK");
                affectsStockNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(affectsStockNode);

                XmlNode forPayrollNode = doc.CreateElement("FORPAYROLL");
                forPayrollNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(forPayrollNode);

                XmlNode interestOnBillwiseNode = doc.CreateElement("INTERESTONBILLWISE");
                interestOnBillwiseNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(interestOnBillwiseNode);

                XmlNode overrideInterestNode = doc.CreateElement("OVERRIDEINTEREST");
                overrideInterestNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(overrideInterestNode);

                XmlNode overrideAdvInterestNode = doc.CreateElement("OVERRIDEADVINTEREST");
                overrideAdvInterestNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(overrideAdvInterestNode);


                XmlNode useForVatNode = doc.CreateElement("USEFORVAT");
                useForVatNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(useForVatNode);

                XmlNode ignoreTdsExemptNode = doc.CreateElement("IGNORETDSEXEMPT");
                ignoreTdsExemptNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(ignoreTdsExemptNode);

                XmlNode isTcsApplicableNode = doc.CreateElement("ISTCSAPPLICABLE");
                isTcsApplicableNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isTcsApplicableNode);

                XmlNode isTdsApplicableNode = doc.CreateElement("ISTDSAPPLICABLE");
                isTdsApplicableNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isTdsApplicableNode);

                XmlNode isFbtApplicableNode = doc.CreateElement("ISFBTAPPLICABLE");
                isFbtApplicableNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isFbtApplicableNode);

                XmlNode isGstApplicable = doc.CreateElement("ISGSTAPPLICABLE");
                isGstApplicable.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isGstApplicable);

                XmlNode isexciseApplicable = doc.CreateElement("ISEXCISEAPPLICABLE");
                isexciseApplicable.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isexciseApplicable);

                XmlNode showInPayslipNode = doc.CreateElement("SHOWINPAYSLIP");
                showInPayslipNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(showInPayslipNode);

                XmlNode useForGratuityNode = doc.CreateElement("USEFORGRATUITY");
                useForGratuityNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(useForGratuityNode);


                XmlNode forServiceTaxNode = doc.CreateElement("FORSERVICETAX");
                forServiceTaxNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(forServiceTaxNode);


                XmlNode isInputCreditNode = doc.CreateElement("ISINPUTCREDIT");
                isInputCreditNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isInputCreditNode);


                XmlNode isExemptedNode = doc.CreateElement("ISEXEMPTED");
                isExemptedNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isExemptedNode);


                XmlNode isAbatementApplicableNode = doc.CreateElement("ISABATEMENTAPPLICABLE");
                isAbatementApplicableNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(isAbatementApplicableNode);


                XmlNode tdsDeducteeIsSpecialRateNode = doc.CreateElement("TDSDEDUCTEEISSPECIALRATE");
                tdsDeducteeIsSpecialRateNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(tdsDeducteeIsSpecialRateNode);


                XmlNode auditedNode = doc.CreateElement("AUDITED");
                auditedNode.AppendChild(doc.CreateTextNode("No"));
                ledgerNode.AppendChild(auditedNode);


                XmlNode sortPositionNode = doc.CreateElement("SORTPOSITION");
                sortPositionNode.AppendChild(doc.CreateTextNode(" 1000"));
                ledgerNode.AppendChild(sortPositionNode);


                XmlNode languageNameListNode = doc.CreateElement("LANGUAGENAME.LIST");
                languageNameListNode.AppendChild(doc.CreateTextNode(""));
                ledgerNode.AppendChild(languageNameListNode);


                XmlNode nameListNode = doc.CreateElement("NAME.LIST");
                XmlAttribute nameListAttribute = doc.CreateAttribute("TYPE");
                nameListAttribute.Value = "String";
                nameListNode.Attributes.Append(nameListAttribute);
                languageNameListNode.AppendChild(nameListNode);

                XmlNode nameListnameNode = doc.CreateElement("NAME");
                nameListnameNode.AppendChild(doc.CreateTextNode(T1));
                nameListNode.AppendChild(nameListnameNode);

                XmlNode languageIdNode = doc.CreateElement("LANGUAGEID");
                languageIdNode.AppendChild(doc.CreateTextNode(" 1033"));
                languageNameListNode.AppendChild(languageIdNode);
            }

            Dt1.Dispose();

            doc.Save(String.IsNullOrEmpty(value)? "C:\\Tally" : value + "\\m-" + input_date);
        }
        #endregion

        #region Export all the Transactions in Tally XML Format
        private void Export_Transactions(String input_date)
        {
            XmlDocument doc = new XmlDocument();
            XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.AppendChild(docNode);

            XmlNode envelopeNode = doc.CreateElement("ENVELOPE");   //ENVELOPE
            doc.AppendChild(envelopeNode);

            XmlNode headerNode = doc.CreateElement("HEADER");       //HEADER
            envelopeNode.AppendChild(headerNode);

            XmlNode tallyRequestNode = doc.CreateElement("TALLYREQUEST");   //TALLY REQUEST
            tallyRequestNode.AppendChild(doc.CreateTextNode("Import Data"));
            headerNode.AppendChild(tallyRequestNode);


            XmlNode bodyNode = doc.CreateElement("BODY");           //BODY
            envelopeNode.AppendChild(bodyNode);

            XmlNode importDataNode = doc.CreateElement("IMPORTDATA");       //IMPORT DATA
            bodyNode.AppendChild(importDataNode);

            XmlNode requestDescNode = doc.CreateElement("REQUESTDESC");     //REQUEST DESC
            importDataNode.AppendChild(requestDescNode);

            XmlNode reportNameNode = doc.CreateElement("REPORTNAME");       //REPORT NAME
            reportNameNode.AppendChild(doc.CreateTextNode("All Masters"));              //Name of the Report
            requestDescNode.AppendChild(reportNameNode);

            XmlNode staticVariablesNode = doc.CreateElement("STATICVARIABLES");
            requestDescNode.AppendChild(staticVariablesNode);

            XmlNode svCurrentCompnayNode = doc.CreateElement("SVCURRENTCOMPANY");       //SVCURRENTCOMPANY

            svCurrentCompnayNode.AppendChild(doc.CreateTextNode("company_name"));      //Datapull 
            staticVariablesNode.AppendChild(svCurrentCompnayNode);

            XmlNode requestDataNode = doc.CreateElement("REQUESTDATA");
            importDataNode.AppendChild(requestDataNode);

            String LoanDisbursed = @"SELECT CASE 
		WHEN cr = '0'
			THEN 'Cash'
		ELSE cr
		END AS cr
	,dr
	,amount
	,CONVERT(VARCHAR(10), dt, 112) AS dt
	,ck_no
FROM (
	SELECT l.L_Id AS id
		,ISNULL(p.P_Name, '0') AS cr
		,l.L_LoanAmount AS amount
		,l.L_ChequeDate AS dt
		,l.L_ChequeNo AS ck_no
	FROM LoanSanction l
	LEFT JOIN PartyMaster AS p ON l.L_BankId = p.P_Id
	) AS t_cr
	,(
		SELECT l.L_Id AS id
			,p.P_Name AS dr
		FROM LoanSanction l
			,PartyMaster AS p
		WHERE l.L_PartyId = p.P_Id
		) AS t_dr
WHERE t_cr.id = t_dr.id
	AND dt IN ('" + input_date + "');";


            DataTable Dt1 = ObjData.GetDataTable(LoanDisbursed);

            for (int i = 0; i < Dt1.Rows.Count; i++)
            {
                String cr = Dt1.Rows[i]["cr"].ToString();
                String dr = Dt1.Rows[i]["dr"].ToString();
                String amount = Math.Round(Convert.ToDouble(Dt1.Rows[i]["amount"]), 0).ToString();
                String dt = Dt1.Rows[i]["dt"].ToString();
                String ck_no = Dt1.Rows[i]["ck_no"].ToString();

                String guid = input_date + "-DISBURSE-" + cr + "-" + dr + "-" + amount + "-" + i.ToString();            //REMOTEID

                XmlNode tallyMessageNode = doc.CreateElement("TALLYMESSAGE");
                XmlAttribute productAttribute = doc.CreateAttribute("xmlns:UDF");
                productAttribute.Value = "TallyUDF";
                tallyMessageNode.Attributes.Append(productAttribute);
                requestDataNode.AppendChild(tallyMessageNode);

                XmlNode voucherNode = doc.CreateElement("VOUCHER");
                XmlAttribute nameAttribute = doc.CreateAttribute("REMOTEID");
                nameAttribute.Value = guid;
                XmlAttribute nameAttribute1 = doc.CreateAttribute("VCHTYPE");
                nameAttribute1.Value = "Payment";                                              //Data Pull
                XmlAttribute nameAttribute2 = doc.CreateAttribute("ACTION");
                nameAttribute2.Value = "Create";
                voucherNode.Attributes.Append(nameAttribute);
                voucherNode.Attributes.Append(nameAttribute1);
                voucherNode.Attributes.Append(nameAttribute2);
                tallyMessageNode.AppendChild(voucherNode);

                XmlNode dateNode = doc.CreateElement("DATE");
                dateNode.AppendChild(doc.CreateTextNode(dt));                           //Data Pull
                voucherNode.AppendChild(dateNode);


                XmlNode guidNode = doc.CreateElement("GUID");
                guidNode.AppendChild(doc.CreateTextNode(guid));
                voucherNode.AppendChild(guidNode);

                XmlNode narrationNode = doc.CreateElement("NARRATION");
                narrationNode.AppendChild(doc.CreateTextNode("Loan Disbursed to " + dr + " of amount " + amount + (String.IsNullOrEmpty(ck_no)? "": " Check no:" + ck_no)));               //Data Pull
                voucherNode.AppendChild(narrationNode);


                XmlNode voucherTypeNameNode = doc.CreateElement("VOUCHERTYPENAME");     //Data Pull
                voucherTypeNameNode.AppendChild(doc.CreateTextNode("Payment"));
                voucherNode.AppendChild(voucherTypeNameNode);

                XmlNode partyLedgerNameNode = doc.CreateElement("PARTYLEDGERNAME");
                partyLedgerNameNode.AppendChild(doc.CreateTextNode(dr));
                voucherNode.AppendChild(partyLedgerNameNode);

                XmlNode fbtPaymentTypeNode = doc.CreateElement("FBTPAYMENTTYPE");
                fbtPaymentTypeNode.AppendChild(doc.CreateTextNode("Default"));
                voucherNode.AppendChild(fbtPaymentTypeNode);

                XmlNode diffactualQtyNode = doc.CreateElement("DIFFACTUALQTY");
                diffactualQtyNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(diffactualQtyNode);

                XmlNode auditedNode = doc.CreateElement("AUDITED");
                auditedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(auditedNode);

                XmlNode forJobCostingNode = doc.CreateElement("FORJOBCOSTING");
                forJobCostingNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(forJobCostingNode);

                XmlNode isOptionalNode = doc.CreateElement("ISOPTIONAL");
                isOptionalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isOptionalNode);

                XmlNode effectivedateNode = doc.CreateElement("EFFECTIVEDATE");
                effectivedateNode.AppendChild(doc.CreateTextNode(dt));
                voucherNode.AppendChild(effectivedateNode);

                XmlNode useForInterestNode = doc.CreateElement("USEFORINTEREST");
                useForInterestNode.AppendChild(doc.CreateTextNode("NO"));
                voucherNode.AppendChild(useForInterestNode);

                XmlNode useforgainlossNode = doc.CreateElement("USEFORGAINLOSS");
                useforgainlossNode.AppendChild(doc.CreateTextNode("NO"));
                voucherNode.AppendChild(useforgainlossNode);

                XmlNode useforgodowntransferNode = doc.CreateElement("USEFORGODOWNTRANSFER");
                useforgodowntransferNode.AppendChild(doc.CreateTextNode("NO"));
                voucherNode.AppendChild(useforgodowntransferNode);

                XmlNode useforcompoundNode = doc.CreateElement("USEFORCOMPOUND");
                useforcompoundNode.AppendChild(doc.CreateTextNode("NO"));
                voucherNode.AppendChild(useforcompoundNode);

                XmlNode alteridNode = doc.CreateElement("ALTERID");
                alteridNode.AppendChild(doc.CreateTextNode(i.ToString()));
                voucherNode.AppendChild(alteridNode);

                XmlNode exciseopeningNode = doc.CreateElement("EXCISEOPENING");
                exciseopeningNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(exciseopeningNode);

                XmlNode useforfinalproductionNode = doc.CreateElement("USEFORFINALPRODUCTION");
                useforfinalproductionNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforfinalproductionNode);

                XmlNode iscancelledNode = doc.CreateElement("ISCANCELLED");
                iscancelledNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscancelledNode);

                XmlNode hascashflowNode = doc.CreateElement("HASCASHFLOW");
                hascashflowNode.AppendChild(doc.CreateTextNode("Yes"));
                voucherNode.AppendChild(hascashflowNode);

                XmlNode ispostdatedNode = doc.CreateElement("ISPOSTDATED");
                ispostdatedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(ispostdatedNode);

                XmlNode usetrackingnumberNode = doc.CreateElement("USETRACKINGNUMBER");
                usetrackingnumberNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(usetrackingnumberNode);

                XmlNode isinvoiceNode = doc.CreateElement("ISINVOICE");
                isinvoiceNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isinvoiceNode);

                XmlNode mfgjournalNode = doc.CreateElement("MFGJOURNAL");
                mfgjournalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(mfgjournalNode);

                XmlNode hasdiscountsNode = doc.CreateElement("HASDISCOUNTS");
                hasdiscountsNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hasdiscountsNode);

                XmlNode aspayslipNode = doc.CreateElement("ASPAYSLIP");
                aspayslipNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(aspayslipNode);

                XmlNode iscostcentreNode = doc.CreateElement("ISCOSTCENTRE");
                iscostcentreNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscostcentreNode);

                XmlNode isdeletedNode = doc.CreateElement("ISDELETED");
                isdeletedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isdeletedNode);

                XmlNode asoriginalNode = doc.CreateElement("ASORIGINAL");
                asoriginalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(asoriginalNode);

                XmlNode allledgerentriesNode = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode);

                //AllLEDGERENTRIES child nodes
                XmlNode ledgernameNode = doc.CreateElement("LEDGERNAME");
                ledgernameNode.AppendChild(doc.CreateTextNode(dr));
                allledgerentriesNode.AppendChild(ledgernameNode);

                XmlNode isdeemedpositiveNode = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositiveNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(isdeemedpositiveNode);

                XmlNode ledgerfromitemNode = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitemNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(ledgerfromitemNode);

                XmlNode removezeroentriesNode = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentriesNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(removezeroentriesNode);

                XmlNode ispartyledgerNode = doc.CreateElement("ISPARTYLEDGER");
                ispartyledgerNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(ispartyledgerNode);

                XmlNode amountNode = doc.CreateElement("AMOUNT");
                amountNode.AppendChild(doc.CreateTextNode("-" + amount));
                allledgerentriesNode.AppendChild(amountNode);

                XmlNode allledgerentriesNode1 = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode1);

                //AllLEDGERENTRIES1 child nodes
                XmlNode ledgername1Node = doc.CreateElement("LEDGERNAME");
                ledgername1Node.AppendChild(doc.CreateTextNode(cr));
                allledgerentriesNode1.AppendChild(ledgername1Node);

                XmlNode isdeemedpositive1Node = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositive1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(isdeemedpositive1Node);

                XmlNode ledgerfromitem1Node = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitem1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(ledgerfromitem1Node);

                XmlNode removezeroentries1Node = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentries1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(removezeroentries1Node);

                XmlNode ispartyledger1Node = doc.CreateElement("ISPARTYLEDGER");
                ispartyledger1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(ispartyledger1Node);

                XmlNode amount1Node = doc.CreateElement("AMOUNT");
                amount1Node.AppendChild(doc.CreateTextNode(amount));
                allledgerentriesNode1.AppendChild(amount1Node);

            }

            String query_EMI = @"SELECT cr
	,CASE 
		WHEN dr = '0'
			THEN 'Cash'
		ELSE dr
		END AS dr
	,amount
	,CONVERT(VARCHAR(10), dt, 112) AS dt
	,eno
	,ck_amt
	,ck_no
	, party_bank
	, party_bank_addr
FROM (
	SELECT emi.C_Id AS id
		,party.P_Name cr
		,emi.C_AmountRec AS amount
		,emi.C_Date AS dt
		,EMI.C_EMINo AS eno
		,cks.S_ChequeDate AS ck_amt
		,cks.S_ChequeNo   AS ck_no
		,sanc.L_PartyBank AS party_bank
		,sanc.L_PartyBankAdd AS party_bank_addr
	FROM 
		LoanSanction sanc
		,LoanApplication app
		,PartyMaster party
		, EMIReceived emi LEFT JOIN ChequeDetails as cks ON C_SId = cks.S_Id AND C_EMINo = cks.S_SNo
	WHERE emi.C_SId = sanc.L_Id
		AND sanc.L_ApplicationId = app.L_Id
		AND app.L_PartyId = party.P_id
	) AS t_cr
	,(
		SELECT emi.C_Id AS id
			,ISNULL(party.P_Name, '0') AS dr
		FROM EMIReceived emi
		LEFT JOIN PartyMaster party ON emi.C_BankId = party.P_Id
		) AS t_dr
WHERE t_cr.id = t_dr.id
	AND dt IN ('" + input_date + "');";

            DataTable Dt2 = ObjData.GetDataTable(query_EMI);

            for (int i = 0; i < Dt2.Rows.Count; i++)
            {
                String cr = Dt2.Rows[i]["cr"].ToString();
                String dr = Dt2.Rows[i]["dr"].ToString();
                String amount = Math.Round(Convert.ToDouble(Dt2.Rows[i]["amount"]), 0).ToString();
                String dt = Dt2.Rows[i]["dt"].ToString();
                String eno = Dt2.Rows[i]["eno"].ToString();
                String ck_amt = Dt2.Rows[i]["ck_amt"].ToString();
                String ck_no = Dt2.Rows[i]["ck_no"].ToString();
                String party_bank = Dt2.Rows[i]["party_bank"].ToString();
                String party_bank_addr = Dt2.Rows[i]["party_bank_addr"].ToString();

                String narration;

                if (dr.Equals("Cash"))
                {
                    narration = "EMI No. " + eno + " Received from " + cr + " in Cash";
                }
                else
                {
                    narration = "EMI No. " + eno + " Received from " + cr + " Payment Details: " + ck_amt + " " + ck_no + " " + party_bank + " " + party_bank_addr;
                }

                String guid = input_date + "-EMI_RECEIVED-" + cr + "-" + dr + "-" + amount + "-" + i.ToString();

                XmlNode tallyMessageNode = doc.CreateElement("TALLYMESSAGE");
                XmlAttribute productAttribute = doc.CreateAttribute("xmlns:UDF");
                productAttribute.Value = "TallyUDF";
                tallyMessageNode.Attributes.Append(productAttribute);
                requestDataNode.AppendChild(tallyMessageNode);

                XmlNode voucherNode = doc.CreateElement("VOUCHER");
                XmlAttribute nameAttribute = doc.CreateAttribute("REMOTEID");
                nameAttribute.Value = guid;
                XmlAttribute nameAttribute1 = doc.CreateAttribute("VCHTYPE");
                nameAttribute1.Value = "Receipt";                                              //Data Pull
                XmlAttribute nameAttribute2 = doc.CreateAttribute("ACTION");
                nameAttribute2.Value = "Create";
                voucherNode.Attributes.Append(nameAttribute);
                voucherNode.Attributes.Append(nameAttribute1);
                voucherNode.Attributes.Append(nameAttribute2);
                tallyMessageNode.AppendChild(voucherNode);

                XmlNode dateNode = doc.CreateElement("DATE");
                dateNode.AppendChild(doc.CreateTextNode(dt));                           //Data Pull
                voucherNode.AppendChild(dateNode);


                XmlNode guidNode = doc.CreateElement("GUID");
                guidNode.AppendChild(doc.CreateTextNode(guid));
                voucherNode.AppendChild(guidNode);

                XmlNode narrationNode = doc.CreateElement("NARRATION");
                narrationNode.AppendChild(doc.CreateTextNode(narration));                  //Data Pull
                voucherNode.AppendChild(narrationNode);


                XmlNode voucherTypeNameNode = doc.CreateElement("VOUCHERTYPENAME");     //Data Pull
                voucherTypeNameNode.AppendChild(doc.CreateTextNode("Receipt"));
                voucherNode.AppendChild(voucherTypeNameNode);

                XmlNode partyLedgerNameNode = doc.CreateElement("PARTYLEDGERNAME");
                partyLedgerNameNode.AppendChild(doc.CreateTextNode(cr));
                voucherNode.AppendChild(partyLedgerNameNode);

                XmlNode fbtPaymentTypeNode = doc.CreateElement("FBTPAYMENTTYPE");
                fbtPaymentTypeNode.AppendChild(doc.CreateTextNode("Default"));
                voucherNode.AppendChild(fbtPaymentTypeNode);

                XmlNode diffactualQtyNode = doc.CreateElement("DIFFACTUALQTY");
                diffactualQtyNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(diffactualQtyNode);

                XmlNode auditedNode = doc.CreateElement("AUDITED");
                auditedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(auditedNode);

                XmlNode forJobCostingNode = doc.CreateElement("FORJOBCOSTING");
                forJobCostingNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(forJobCostingNode);

                XmlNode isOptionalNode = doc.CreateElement("ISOPTIONAL");
                isOptionalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isOptionalNode);

                XmlNode effectivedateNode = doc.CreateElement("EFFECTIVEDATE");
                effectivedateNode.AppendChild(doc.CreateTextNode(dt));
                voucherNode.AppendChild(effectivedateNode);

                XmlNode useForInterestNode = doc.CreateElement("USEFORINTEREST");
                useForInterestNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useForInterestNode);

                XmlNode useforgainlossNode = doc.CreateElement("USEFORGAINLOSS");
                useforgainlossNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgainlossNode);

                XmlNode useforgodowntransferNode = doc.CreateElement("USEFORGODOWNTRANSFER");
                useforgodowntransferNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgodowntransferNode);

                XmlNode useforcompoundNode = doc.CreateElement("USEFORCOMPOUND");
                useforcompoundNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforcompoundNode);

                XmlNode alteridNode = doc.CreateElement("ALTERID");
                alteridNode.AppendChild(doc.CreateTextNode(i.ToString()));
                voucherNode.AppendChild(alteridNode);

                XmlNode exciseopeningNode = doc.CreateElement("EXCISEOPENING");
                exciseopeningNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(exciseopeningNode);

                XmlNode useforfinalproductionNode = doc.CreateElement("USEFORFINALPRODUCTION");
                useforfinalproductionNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforfinalproductionNode);

                XmlNode iscancelledNode = doc.CreateElement("ISCANCELLED");
                iscancelledNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscancelledNode);

                XmlNode hascashflowNode = doc.CreateElement("HASCASHFLOW");
                hascashflowNode.AppendChild(doc.CreateTextNode("Yes"));
                voucherNode.AppendChild(hascashflowNode);

                XmlNode ispostdatedNode = doc.CreateElement("ISPOSTDATED");
                ispostdatedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(ispostdatedNode);

                XmlNode usetrackingnumberNode = doc.CreateElement("USETRACKINGNUMBER");
                usetrackingnumberNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(usetrackingnumberNode);

                XmlNode isinvoiceNode = doc.CreateElement("ISINVOICE");
                isinvoiceNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isinvoiceNode);

                XmlNode mfgjournalNode = doc.CreateElement("MFGJOURNAL");
                mfgjournalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(mfgjournalNode);

                XmlNode hasdiscountsNode = doc.CreateElement("HASDISCOUNTS");
                hasdiscountsNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hasdiscountsNode);

                XmlNode aspayslipNode = doc.CreateElement("ASPAYSLIP");
                aspayslipNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(aspayslipNode);

                XmlNode iscostcentreNode = doc.CreateElement("ISCOSTCENTRE");
                iscostcentreNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscostcentreNode);

                XmlNode isdeletedNode = doc.CreateElement("ISDELETED");
                isdeletedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isdeletedNode);

                XmlNode asoriginalNode = doc.CreateElement("ASORIGINAL");
                asoriginalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(asoriginalNode);

                XmlNode allledgerentriesNode = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode);

                //AllLEDGERENTRIES child nodes
                XmlNode ledgernameNode = doc.CreateElement("LEDGERNAME");
                ledgernameNode.AppendChild(doc.CreateTextNode(cr));
                allledgerentriesNode.AppendChild(ledgernameNode);

                XmlNode isdeemedpositiveNode = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositiveNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(isdeemedpositiveNode);

                XmlNode ledgerfromitemNode = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitemNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(ledgerfromitemNode);

                XmlNode removezeroentriesNode = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentriesNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(removezeroentriesNode);

                XmlNode ispartyledgerNode = doc.CreateElement("ISPARTYLEDGER");
                ispartyledgerNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(ispartyledgerNode);

                XmlNode amountNode = doc.CreateElement("AMOUNT");
                amountNode.AppendChild(doc.CreateTextNode(amount));
                allledgerentriesNode.AppendChild(amountNode);

                XmlNode allledgerentriesNode1 = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode1);

                //AllLEDGERENTRIES1 child nodes
                XmlNode ledgername1Node = doc.CreateElement("LEDGERNAME");
                ledgername1Node.AppendChild(doc.CreateTextNode(dr));
                allledgerentriesNode1.AppendChild(ledgername1Node);

                XmlNode isdeemedpositive1Node = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositive1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(isdeemedpositive1Node);

                XmlNode ledgerfromitem1Node = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitem1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(ledgerfromitem1Node);

                XmlNode removezeroentries1Node = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentries1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(removezeroentries1Node);

                XmlNode ispartyledger1Node = doc.CreateElement("ISPARTYLEDGER");
                ispartyledger1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(ispartyledger1Node);

                XmlNode amount1Node = doc.CreateElement("AMOUNT");
                amount1Node.AppendChild(doc.CreateTextNode("-" + amount));
                allledgerentriesNode1.AppendChild(amount1Node);
            }

            String ForeClosure_Interest = @"SELECT 'Interest A/c' AS cr
	,party.P_Name AS dr
	,pmts.F_InterestAmount amount
	,CONVERT(VARCHAR(10), pmts.F_Date, 112) AS dt
FROM AccountForeClosure AS pmts
	,LoanSanction AS loan
	,PartyMaster AS party
WHERE pmts.F_InterestAmount <> 0
	AND pmts.F_SId = loan.L_Id
	AND loan.L_PartyId = party.P_Id
	AND pmts.F_Date IN ('" + input_date + "');";

            DataTable Dt3 = ObjData.GetDataTable(ForeClosure_Interest);

            for (int i = 0; i < Dt3.Rows.Count; i++)
            {
                String cr = Dt3.Rows[i]["cr"].ToString();
                String dr = Dt3.Rows[i]["dr"].ToString();
                String amount = Math.Round(Convert.ToDouble(Dt3.Rows[i]["amount"]), 0).ToString();
                String dt = Dt3.Rows[i]["dt"].ToString();

                String guid = input_date + "-FORECLOSURE_INTEREST-" + cr + "-" + dr + "-" + amount + "-" + i.ToString();

                XmlNode tallyMessageNode = doc.CreateElement("TALLYMESSAGE");
                XmlAttribute productAttribute = doc.CreateAttribute("xmlns:UDF");
                productAttribute.Value = "TallyUDF";
                tallyMessageNode.Attributes.Append(productAttribute);
                requestDataNode.AppendChild(tallyMessageNode);

                XmlNode voucherNode = doc.CreateElement("VOUCHER");
                XmlAttribute nameAttribute = doc.CreateAttribute("REMOTEID");
                nameAttribute.Value = guid;
                XmlAttribute nameAttribute1 = doc.CreateAttribute("VCHTYPE");
                nameAttribute1.Value = "Journal";                                              //Data Pull
                XmlAttribute nameAttribute2 = doc.CreateAttribute("ACTION");
                nameAttribute2.Value = "Create";
                voucherNode.Attributes.Append(nameAttribute);
                voucherNode.Attributes.Append(nameAttribute1);
                voucherNode.Attributes.Append(nameAttribute2);
                tallyMessageNode.AppendChild(voucherNode);

                XmlNode dateNode = doc.CreateElement("DATE");
                dateNode.AppendChild(doc.CreateTextNode(dt));                           //Data Pull
                voucherNode.AppendChild(dateNode);


                XmlNode guidNode = doc.CreateElement("GUID");
                guidNode.AppendChild(doc.CreateTextNode(guid));
                voucherNode.AppendChild(guidNode);

                XmlNode narrationNode = doc.CreateElement("NARRATION");
                narrationNode.AppendChild(doc.CreateTextNode("ForeClosure Interest " + amount + " by " + dr));                  //Data Pull
                voucherNode.AppendChild(narrationNode);


                XmlNode voucherTypeNameNode = doc.CreateElement("VOUCHERTYPENAME");     //Data Pull
                voucherTypeNameNode.AppendChild(doc.CreateTextNode("Journal"));
                voucherNode.AppendChild(voucherTypeNameNode);

                XmlNode partyLedgerNameNode = doc.CreateElement("PARTYLEDGERNAME");
                partyLedgerNameNode.AppendChild(doc.CreateTextNode(dr));
                voucherNode.AppendChild(partyLedgerNameNode);

                XmlNode fbtPaymentTypeNode = doc.CreateElement("FBTPAYMENTTYPE");
                fbtPaymentTypeNode.AppendChild(doc.CreateTextNode("Default"));
                voucherNode.AppendChild(fbtPaymentTypeNode);

                XmlNode diffactualQtyNode = doc.CreateElement("DIFFACTUALQTY");
                diffactualQtyNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(diffactualQtyNode);

                XmlNode auditedNode = doc.CreateElement("AUDITED");
                auditedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(auditedNode);

                XmlNode forJobCostingNode = doc.CreateElement("FORJOBCOSTING");
                forJobCostingNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(forJobCostingNode);

                XmlNode isOptionalNode = doc.CreateElement("ISOPTIONAL");
                isOptionalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isOptionalNode);

                XmlNode effectivedateNode = doc.CreateElement("EFFECTIVEDATE");
                effectivedateNode.AppendChild(doc.CreateTextNode(dt));
                voucherNode.AppendChild(effectivedateNode);

                XmlNode useForInterestNode = doc.CreateElement("USEFORINTEREST");
                useForInterestNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useForInterestNode);

                XmlNode useforgainlossNode = doc.CreateElement("USEFORGAINLOSS");
                useforgainlossNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgainlossNode);

                XmlNode useforgodowntransferNode = doc.CreateElement("USEFORGODOWNTRANSFER");
                useforgodowntransferNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgodowntransferNode);

                XmlNode useforcompoundNode = doc.CreateElement("USEFORCOMPOUND");
                useforcompoundNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforcompoundNode);

                XmlNode alteridNode = doc.CreateElement("ALTERID");
                alteridNode.AppendChild(doc.CreateTextNode(i.ToString()));
                voucherNode.AppendChild(alteridNode);

                XmlNode exciseopeningNode = doc.CreateElement("EXCISEOPENING");
                exciseopeningNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(exciseopeningNode);

                XmlNode useforfinalproductionNode = doc.CreateElement("USEFORFINALPRODUCTION");
                useforfinalproductionNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforfinalproductionNode);

                XmlNode iscancelledNode = doc.CreateElement("ISCANCELLED");
                iscancelledNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscancelledNode);

                XmlNode hascashflowNode = doc.CreateElement("HASCASHFLOW");
                hascashflowNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hascashflowNode);

                XmlNode ispostdatedNode = doc.CreateElement("ISPOSTDATED");
                ispostdatedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(ispostdatedNode);

                XmlNode usetrackingnumberNode = doc.CreateElement("USETRACKINGNUMBER");
                usetrackingnumberNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(usetrackingnumberNode);

                XmlNode isinvoiceNode = doc.CreateElement("ISINVOICE");
                isinvoiceNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isinvoiceNode);

                XmlNode mfgjournalNode = doc.CreateElement("MFGJOURNAL");
                mfgjournalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(mfgjournalNode);

                XmlNode hasdiscountsNode = doc.CreateElement("HASDISCOUNTS");
                hasdiscountsNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hasdiscountsNode);

                XmlNode aspayslipNode = doc.CreateElement("ASPAYSLIP");
                aspayslipNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(aspayslipNode);

                XmlNode iscostcentreNode = doc.CreateElement("ISCOSTCENTRE");
                iscostcentreNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscostcentreNode);

                XmlNode isdeletedNode = doc.CreateElement("ISDELETED");
                isdeletedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isdeletedNode);

                XmlNode asoriginalNode = doc.CreateElement("ASORIGINAL");
                asoriginalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(asoriginalNode);

                XmlNode allledgerentriesNode = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode);

                //AllLEDGERENTRIES child nodes
                XmlNode ledgernameNode = doc.CreateElement("LEDGERNAME");
                ledgernameNode.AppendChild(doc.CreateTextNode(dr));
                allledgerentriesNode.AppendChild(ledgernameNode);

                XmlNode isdeemedpositiveNode = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositiveNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(isdeemedpositiveNode);

                XmlNode ledgerfromitemNode = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitemNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(ledgerfromitemNode);

                XmlNode removezeroentriesNode = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentriesNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(removezeroentriesNode);

                XmlNode ispartyledgerNode = doc.CreateElement("ISPARTYLEDGER");
                ispartyledgerNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(ispartyledgerNode);

                XmlNode amountNode = doc.CreateElement("AMOUNT");
                amountNode.AppendChild(doc.CreateTextNode("-" + amount));
                allledgerentriesNode.AppendChild(amountNode);

                XmlNode allledgerentriesNode1 = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode1);

                //AllLEDGERENTRIES1 child nodes
                XmlNode ledgername1Node = doc.CreateElement("LEDGERNAME");
                ledgername1Node.AppendChild(doc.CreateTextNode(cr));
                allledgerentriesNode1.AppendChild(ledgername1Node);

                XmlNode isdeemedpositive1Node = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositive1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(isdeemedpositive1Node);

                XmlNode ledgerfromitem1Node = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitem1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(ledgerfromitem1Node);

                XmlNode removezeroentries1Node = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentries1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(removezeroentries1Node);

                XmlNode ispartyledger1Node = doc.CreateElement("ISPARTYLEDGER");
                ispartyledger1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(ispartyledger1Node);

                XmlNode amount1Node = doc.CreateElement("AMOUNT");
                amount1Node.AppendChild(doc.CreateTextNode(amount));
                allledgerentriesNode1.AppendChild(amount1Node);
            }


            String query_InterestAccrual = @"SELECT 'Interest A/c' AS cr
	,party.P_Name AS dr
	,(pmts.InterestPer / 100) * emi.C_AmountRec AS amount
	,CONVERT(VARCHAR(10), emi.C_Date, 112) AS dt
FROM EmiReceived emi
	,TmpDataCal AS pmts
	,LoanSanction AS loan
	,PartyMaster AS party
WHERE emi.C_SId = loan.L_Id
	AND pmts.Id = loan.L_Id
	AND loan.L_PartyId = party.P_Id
	AND emi.C_EMINo = pmts.EMINo
	AND emi.C_Date IN ('" + input_date + "');";

            DataTable Dt4 = ObjData.GetDataTable(query_InterestAccrual);

            for (int i = 0; i < Dt4.Rows.Count; i++)
            {
                String cr = Dt4.Rows[i]["cr"].ToString();
                String dr = Dt4.Rows[i]["dr"].ToString();
                String amount = Math.Round(Convert.ToDouble(Dt4.Rows[i]["amount"]), 0).ToString();
                String dt = Dt4.Rows[i]["dt"].ToString();

                String guid = input_date + "-INTEREST_ACCRUAL-" + cr + "-" + dr + "-" + amount + "-" + i.ToString();

                XmlNode tallyMessageNode = doc.CreateElement("TALLYMESSAGE");
                XmlAttribute productAttribute = doc.CreateAttribute("xmlns:UDF");
                productAttribute.Value = "TallyUDF";
                tallyMessageNode.Attributes.Append(productAttribute);
                requestDataNode.AppendChild(tallyMessageNode);

                XmlNode voucherNode = doc.CreateElement("VOUCHER");
                XmlAttribute nameAttribute = doc.CreateAttribute("REMOTEID");
                nameAttribute.Value = guid;
                XmlAttribute nameAttribute1 = doc.CreateAttribute("VCHTYPE");
                nameAttribute1.Value = "Journal";                                              //Data Pull
                XmlAttribute nameAttribute2 = doc.CreateAttribute("ACTION");
                nameAttribute2.Value = "Create";
                voucherNode.Attributes.Append(nameAttribute);
                voucherNode.Attributes.Append(nameAttribute1);
                voucherNode.Attributes.Append(nameAttribute2);
                tallyMessageNode.AppendChild(voucherNode);

                XmlNode dateNode = doc.CreateElement("DATE");
                dateNode.AppendChild(doc.CreateTextNode(dt));                           //Data Pull
                voucherNode.AppendChild(dateNode);


                XmlNode guidNode = doc.CreateElement("GUID");
                guidNode.AppendChild(doc.CreateTextNode(guid));
                voucherNode.AppendChild(guidNode);

                XmlNode narrationNode = doc.CreateElement("NARRATION");
                narrationNode.AppendChild(doc.CreateTextNode("Interest Accrual of " + amount + " by " + dr));                  //Data Pull
                voucherNode.AppendChild(narrationNode);


                XmlNode voucherTypeNameNode = doc.CreateElement("VOUCHERTYPENAME");     //Data Pull
                voucherTypeNameNode.AppendChild(doc.CreateTextNode("Journal"));
                voucherNode.AppendChild(voucherTypeNameNode);

                XmlNode partyLedgerNameNode = doc.CreateElement("PARTYLEDGERNAME");
                partyLedgerNameNode.AppendChild(doc.CreateTextNode(dr));
                voucherNode.AppendChild(partyLedgerNameNode);

                XmlNode fbtPaymentTypeNode = doc.CreateElement("FBTPAYMENTTYPE");
                fbtPaymentTypeNode.AppendChild(doc.CreateTextNode("Default"));
                voucherNode.AppendChild(fbtPaymentTypeNode);

                XmlNode diffactualQtyNode = doc.CreateElement("DIFFACTUALQTY");
                diffactualQtyNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(diffactualQtyNode);

                XmlNode auditedNode = doc.CreateElement("AUDITED");
                auditedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(auditedNode);

                XmlNode forJobCostingNode = doc.CreateElement("FORJOBCOSTING");
                forJobCostingNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(forJobCostingNode);

                XmlNode isOptionalNode = doc.CreateElement("ISOPTIONAL");
                isOptionalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isOptionalNode);

                XmlNode effectivedateNode = doc.CreateElement("EFFECTIVEDATE");
                effectivedateNode.AppendChild(doc.CreateTextNode(dt));
                voucherNode.AppendChild(effectivedateNode);

                XmlNode useForInterestNode = doc.CreateElement("USEFORINTEREST");
                useForInterestNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useForInterestNode);

                XmlNode useforgainlossNode = doc.CreateElement("USEFORGAINLOSS");
                useforgainlossNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgainlossNode);

                XmlNode useforgodowntransferNode = doc.CreateElement("USEFORGODOWNTRANSFER");
                useforgodowntransferNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgodowntransferNode);

                XmlNode useforcompoundNode = doc.CreateElement("USEFORCOMPOUND");
                useforcompoundNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforcompoundNode);

                XmlNode alteridNode = doc.CreateElement("ALTERID");
                alteridNode.AppendChild(doc.CreateTextNode(i.ToString()));
                voucherNode.AppendChild(alteridNode);

                XmlNode exciseopeningNode = doc.CreateElement("EXCISEOPENING");
                exciseopeningNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(exciseopeningNode);

                XmlNode useforfinalproductionNode = doc.CreateElement("USEFORFINALPRODUCTION");
                useforfinalproductionNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforfinalproductionNode);

                XmlNode iscancelledNode = doc.CreateElement("ISCANCELLED");
                iscancelledNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscancelledNode);

                XmlNode hascashflowNode = doc.CreateElement("HASCASHFLOW");
                hascashflowNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hascashflowNode);

                XmlNode ispostdatedNode = doc.CreateElement("ISPOSTDATED");
                ispostdatedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(ispostdatedNode);

                XmlNode usetrackingnumberNode = doc.CreateElement("USETRACKINGNUMBER");
                usetrackingnumberNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(usetrackingnumberNode);

                XmlNode isinvoiceNode = doc.CreateElement("ISINVOICE");
                isinvoiceNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isinvoiceNode);

                XmlNode mfgjournalNode = doc.CreateElement("MFGJOURNAL");
                mfgjournalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(mfgjournalNode);

                XmlNode hasdiscountsNode = doc.CreateElement("HASDISCOUNTS");
                hasdiscountsNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hasdiscountsNode);

                XmlNode aspayslipNode = doc.CreateElement("ASPAYSLIP");
                aspayslipNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(aspayslipNode);

                XmlNode iscostcentreNode = doc.CreateElement("ISCOSTCENTRE");
                iscostcentreNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscostcentreNode);

                XmlNode isdeletedNode = doc.CreateElement("ISDELETED");
                isdeletedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isdeletedNode);

                XmlNode asoriginalNode = doc.CreateElement("ASORIGINAL");
                asoriginalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(asoriginalNode);

                XmlNode allledgerentriesNode = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode);

                //AllLEDGERENTRIES child nodes
                XmlNode ledgernameNode = doc.CreateElement("LEDGERNAME");
                ledgernameNode.AppendChild(doc.CreateTextNode(dr));
                allledgerentriesNode.AppendChild(ledgernameNode);

                XmlNode isdeemedpositiveNode = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositiveNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(isdeemedpositiveNode);

                XmlNode ledgerfromitemNode = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitemNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(ledgerfromitemNode);

                XmlNode removezeroentriesNode = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentriesNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(removezeroentriesNode);

                XmlNode ispartyledgerNode = doc.CreateElement("ISPARTYLEDGER");
                ispartyledgerNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(ispartyledgerNode);

                XmlNode amountNode = doc.CreateElement("AMOUNT");
                amountNode.AppendChild(doc.CreateTextNode("-" + amount));
                allledgerentriesNode.AppendChild(amountNode);

                XmlNode allledgerentriesNode1 = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode1);

                //AllLEDGERENTRIES1 child nodes
                XmlNode ledgername1Node = doc.CreateElement("LEDGERNAME");
                ledgername1Node.AppendChild(doc.CreateTextNode(cr));
                allledgerentriesNode1.AppendChild(ledgername1Node);

                XmlNode isdeemedpositive1Node = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositive1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(isdeemedpositive1Node);

                XmlNode ledgerfromitem1Node = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitem1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(ledgerfromitem1Node);

                XmlNode removezeroentries1Node = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentries1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(removezeroentries1Node);

                XmlNode ispartyledger1Node = doc.CreateElement("ISPARTYLEDGER");
                ispartyledger1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(ispartyledger1Node);

                XmlNode amount1Node = doc.CreateElement("AMOUNT");
                amount1Node.AppendChild(doc.CreateTextNode(amount));
                allledgerentriesNode1.AppendChild(amount1Node);
            }

            String S5 = "SELECT cr, case when dr ='0' then 'Cash' else dr end as dr, amount, CONVERT(VARCHAR(10), dt, 112) dt  FROM(Select fp.F_Id as id, p.P_Name as cr, (fp.F_Principal + fp.F_InterestAmount + fp.F_Previous) as amount, fp.F_Date as dt from AccountForeClosure as fp, LoanSanction as ls, PartyMaster as p WHERE fp.F_SId = ls.L_Id AND ls.L_PartyId = P.P_Id) as t_cr,(SELECT fp.F_Id as id, ISNULL(party.P_Name,'0') AS dr FROM AccountForeClosure fp LEFT OUTER JOIN PartyMaster party  on fp.F_BankId = party.P_Id) as t_dr WHERE t_cr.id = t_dr.id AND dt in ('" + input_date + "');";
            DataTable Dt5 = ObjData.GetDataTable(S5);

            for (int i = 0; i < Dt5.Rows.Count; i++)
            {
                String cr = Dt5.Rows[i]["cr"].ToString();
                String dr = Dt5.Rows[i]["dr"].ToString();
                String amount = Math.Round(Convert.ToDouble(Dt5.Rows[i]["amount"]), 0).ToString();
                String dt = Dt5.Rows[i]["dt"].ToString();

                String guid = input_date + "-FORECLOSED-" + cr + "-" + dr + "-" + amount + "-" + i.ToString();

                XmlNode tallyMessageNode = doc.CreateElement("TALLYMESSAGE");
                XmlAttribute productAttribute = doc.CreateAttribute("xmlns:UDF");
                productAttribute.Value = "TallyUDF";
                tallyMessageNode.Attributes.Append(productAttribute);
                requestDataNode.AppendChild(tallyMessageNode);

                XmlNode voucherNode = doc.CreateElement("VOUCHER");
                XmlAttribute nameAttribute = doc.CreateAttribute("REMOTEID");
                nameAttribute.Value = guid;
                XmlAttribute nameAttribute1 = doc.CreateAttribute("VCHTYPE");
                nameAttribute1.Value = "Receipt";                                              //Data Pull
                XmlAttribute nameAttribute2 = doc.CreateAttribute("ACTION");
                nameAttribute2.Value = "Create";
                voucherNode.Attributes.Append(nameAttribute);
                voucherNode.Attributes.Append(nameAttribute1);
                voucherNode.Attributes.Append(nameAttribute2);
                tallyMessageNode.AppendChild(voucherNode);

                XmlNode dateNode = doc.CreateElement("DATE");
                dateNode.AppendChild(doc.CreateTextNode(dt));                           //Data Pull
                voucherNode.AppendChild(dateNode);


                XmlNode guidNode = doc.CreateElement("GUID");
                guidNode.AppendChild(doc.CreateTextNode(guid));
                voucherNode.AppendChild(guidNode);

                XmlNode narrationNode = doc.CreateElement("NARRATION");
                narrationNode.AppendChild(doc.CreateTextNode("Foreclosure " + amount + " by " + cr));                  //Data Pull
                voucherNode.AppendChild(narrationNode);


                XmlNode voucherTypeNameNode = doc.CreateElement("VOUCHERTYPENAME");     //Data Pull
                voucherTypeNameNode.AppendChild(doc.CreateTextNode("Receipt"));
                voucherNode.AppendChild(voucherTypeNameNode);

                XmlNode partyLedgerNameNode = doc.CreateElement("PARTYLEDGERNAME");
                partyLedgerNameNode.AppendChild(doc.CreateTextNode(cr));
                voucherNode.AppendChild(partyLedgerNameNode);

                XmlNode fbtPaymentTypeNode = doc.CreateElement("FBTPAYMENTTYPE");
                fbtPaymentTypeNode.AppendChild(doc.CreateTextNode("Default"));
                voucherNode.AppendChild(fbtPaymentTypeNode);

                XmlNode diffactualQtyNode = doc.CreateElement("DIFFACTUALQTY");
                diffactualQtyNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(diffactualQtyNode);

                XmlNode auditedNode = doc.CreateElement("AUDITED");
                auditedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(auditedNode);

                XmlNode forJobCostingNode = doc.CreateElement("FORJOBCOSTING");
                forJobCostingNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(forJobCostingNode);

                XmlNode isOptionalNode = doc.CreateElement("ISOPTIONAL");
                isOptionalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isOptionalNode);

                XmlNode effectivedateNode = doc.CreateElement("EFFECTIVEDATE");
                effectivedateNode.AppendChild(doc.CreateTextNode(dt));
                voucherNode.AppendChild(effectivedateNode);

                XmlNode useForInterestNode = doc.CreateElement("USEFORINTEREST");
                useForInterestNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useForInterestNode);

                XmlNode useforgainlossNode = doc.CreateElement("USEFORGAINLOSS");
                useforgainlossNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgainlossNode);

                XmlNode useforgodowntransferNode = doc.CreateElement("USEFORGODOWNTRANSFER");
                useforgodowntransferNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforgodowntransferNode);

                XmlNode useforcompoundNode = doc.CreateElement("USEFORCOMPOUND");
                useforcompoundNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforcompoundNode);

                XmlNode alteridNode = doc.CreateElement("ALTERID");
                alteridNode.AppendChild(doc.CreateTextNode(i.ToString()));
                voucherNode.AppendChild(alteridNode);

                XmlNode exciseopeningNode = doc.CreateElement("EXCISEOPENING");
                exciseopeningNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(exciseopeningNode);

                XmlNode useforfinalproductionNode = doc.CreateElement("USEFORFINALPRODUCTION");
                useforfinalproductionNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(useforfinalproductionNode);

                XmlNode iscancelledNode = doc.CreateElement("ISCANCELLED");
                iscancelledNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscancelledNode);

                XmlNode hascashflowNode = doc.CreateElement("HASCASHFLOW");
                hascashflowNode.AppendChild(doc.CreateTextNode("Yes"));
                voucherNode.AppendChild(hascashflowNode);

                XmlNode ispostdatedNode = doc.CreateElement("ISPOSTDATED");
                ispostdatedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(ispostdatedNode);

                XmlNode usetrackingnumberNode = doc.CreateElement("USETRACKINGNUMBER");
                usetrackingnumberNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(usetrackingnumberNode);

                XmlNode isinvoiceNode = doc.CreateElement("ISINVOICE");
                isinvoiceNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isinvoiceNode);

                XmlNode mfgjournalNode = doc.CreateElement("MFGJOURNAL");
                mfgjournalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(mfgjournalNode);

                XmlNode hasdiscountsNode = doc.CreateElement("HASDISCOUNTS");
                hasdiscountsNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(hasdiscountsNode);

                XmlNode aspayslipNode = doc.CreateElement("ASPAYSLIP");
                aspayslipNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(aspayslipNode);

                XmlNode iscostcentreNode = doc.CreateElement("ISCOSTCENTRE");
                iscostcentreNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(iscostcentreNode);

                XmlNode isdeletedNode = doc.CreateElement("ISDELETED");
                isdeletedNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(isdeletedNode);

                XmlNode asoriginalNode = doc.CreateElement("ASORIGINAL");
                asoriginalNode.AppendChild(doc.CreateTextNode("No"));
                voucherNode.AppendChild(asoriginalNode);

                XmlNode allledgerentriesNode = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode);

                //AllLEDGERENTRIES child nodes
                XmlNode ledgernameNode = doc.CreateElement("LEDGERNAME");
                ledgernameNode.AppendChild(doc.CreateTextNode(cr));
                allledgerentriesNode.AppendChild(ledgernameNode);

                XmlNode isdeemedpositiveNode = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositiveNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(isdeemedpositiveNode);

                XmlNode ledgerfromitemNode = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitemNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(ledgerfromitemNode);

                XmlNode removezeroentriesNode = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentriesNode.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode.AppendChild(removezeroentriesNode);

                XmlNode ispartyledgerNode = doc.CreateElement("ISPARTYLEDGER");
                ispartyledgerNode.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode.AppendChild(ispartyledgerNode);

                XmlNode amountNode = doc.CreateElement("AMOUNT");
                amountNode.AppendChild(doc.CreateTextNode(amount));
                allledgerentriesNode.AppendChild(amountNode);

                XmlNode allledgerentriesNode1 = doc.CreateElement("ALLLEDGERENTRIES.LIST");
                voucherNode.AppendChild(allledgerentriesNode1);

                //AllLEDGERENTRIES1 child nodes
                XmlNode ledgername1Node = doc.CreateElement("LEDGERNAME");
                ledgername1Node.AppendChild(doc.CreateTextNode(dr));
                allledgerentriesNode1.AppendChild(ledgername1Node);

                XmlNode isdeemedpositive1Node = doc.CreateElement("ISDEEMEDPOSITIVE");
                isdeemedpositive1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(isdeemedpositive1Node);

                XmlNode ledgerfromitem1Node = doc.CreateElement("LEDGERFROMITEM");
                ledgerfromitem1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(ledgerfromitem1Node);

                XmlNode removezeroentries1Node = doc.CreateElement("REMOVEZEROENTRIES");
                removezeroentries1Node.AppendChild(doc.CreateTextNode("No"));
                allledgerentriesNode1.AppendChild(removezeroentries1Node);

                XmlNode ispartyledger1Node = doc.CreateElement("ISPARTYLEDGER");
                ispartyledger1Node.AppendChild(doc.CreateTextNode("Yes"));
                allledgerentriesNode1.AppendChild(ispartyledger1Node);

                XmlNode amount1Node = doc.CreateElement("AMOUNT");
                amount1Node.AppendChild(doc.CreateTextNode("-" + amount));
                allledgerentriesNode1.AppendChild(amount1Node);
            }

            doc.Save(String.IsNullOrEmpty(value) ? "C:\\Tally" : value + "\\v-" + input_date);

            Dt1.Dispose();
            Dt2.Dispose();
            Dt3.Dispose();
            Dt4.Dispose();
            Dt5.Dispose();
        }
        #endregion

        private void form_TallyExport_Load(object sender, EventArgs e)
        {

        }
    }
}

/* 
 * 
 * The default FINANCE.STL file has the following contents:
 * 
 * Data Source=[SERVERNAME]\SQLEXPRESS;Initial Catalog=FinanceManagement;Trusted_Connection=Yes;pooling=false;MultipleActiveResultSets = True;
 * 
 * */
