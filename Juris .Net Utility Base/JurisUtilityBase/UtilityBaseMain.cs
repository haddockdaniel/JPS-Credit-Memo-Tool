using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties
        //152557.82
        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }
        public List<string> invoices = new List<string>();

        public List<CreditMemo> memos = new List<CreditMemo>();

        private string PYear = "";

        private string PNbr = "";

        private string DOrder = "";

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            if (memos != null && memos.Count > 0)
            {
                Cursor.Current = Cursors.WaitCursor;
                toolStripStatusLabel.Text = "Creating Credit Memos...";
                statusStrip.Refresh();
                UpdateStatus("Creating Credit Memos...", 1, memos.Count + 1);
                Application.DoEvents();
                string SQLC = "select max(case when spname='CurAcctPrdYear' then cast(spnbrvalue as varchar(4)) else '' end) as PrdYear, max(Case when spname = 'CurAcctPrdNbr' then case " +
    " when spnbrvalue<9 then '0' + cast(spnbrvalue as varchar(1)) else cast(spnbrvalue as varchar(2)) end  else '' end) as PrdNbr," +
    "max(case when spname='CfgMiscOpts' then substring(sptxtvalue,14,1) else 0 end) as DOrder from sysparam";
                DataSet myRSSysParm = _jurisUtility.RecordsetFromSQL(SQLC);

                DataTable dtSP = myRSSysParm.Tables[0];

                if (dtSP.Rows.Count == 0)
                { MessageBox.Show("Incorrect SysParams"); }
                else
                {
                    foreach (DataRow dr in dtSP.Rows)
                    {
                        PYear = dr["PrdYear"].ToString();
                        PNbr = dr["PrdNbr"].ToString();
                        DOrder = dr["DOrder"].ToString();

                    }
                }
                int counter = 1;
                foreach (CreditMemo cm in memos)
                {
                    cm.LHID = CreateLedgerHistory(cm);
                    toolStripStatusLabel.Text = "Processing Credit Memos...";
                    statusStrip.Refresh();
                    UpdateStatus("Processing Credit Memos...", counter + 1, memos.Count + 1);
                    counter++;
                }

            }
            else
            {
                MessageBox.Show("There were no valid transactions to process - Credit Memo List = 0", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

                    //DataTable d1 = (DataTable)dataGridView1.DataSource;
            
            Cursor.Current = Cursors.Default;
            toolStripStatusLabel.Text = "Utility Completed.";
            statusStrip.Refresh();
            UpdateStatus("Utility Completed.", memos.Count + 1, memos.Count + 1);
            Application.DoEvents();



            string cmt = Application.ProductName.ToString();
            WriteLog("JPS - Credit Memo Utility");

            MessageBox.Show("Credit Memo Creation Completed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.None);
         
        }

        private int CreateLedgerHistory(CreditMemo cm)
        {
            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Creating Ledger History Record...";
            statusStrip.Refresh();
            Application.DoEvents();
            int lastLH = 0;

            string sqlB = "select SpNbrValue from sysparam where spname = 'LastSysNbrLH'";
            DataSet spBatch = _jurisUtility.RecordsetFromSQL(sqlB);
            DataTable dtB = spBatch.Tables[0];
            if (dtB.Rows.Count == 0)
            { MessageBox.Show("Invalid sysparam data - LastSysNbrLH"); }
            else
            {
                foreach (DataRow dr in dtB.Rows)
                {
                    lastLH = Convert.ToInt32(dr["SpNbrValue"].ToString());
                }

            }



            string SQL = "Insert into ledgerhistory( [LHSysNbr] ,[LHMatter] ,[LHBillNbr] ,[LHType] ,[LHDate] ,[LHPrdYear] ,[LHPrdNbr] ,[LHCashAmt]  ,[LHFees] ,[LHCshExp] ,[LHNCshExp] ,[LHSurcharge] ,[LHTaxes1] ,[LHTaxes2] ,[LHTaxes3] ,[LHInterest] ,[LHComment]) VALUES (" +
              (lastLH + 1).ToString() + "," + cm.mat + "," + cm.inv + ",8,'" + DateTime.Today + "', " + PYear + ", " + PNbr + ", 0.00, " + cm.fees * -1 + "," + cm.cashexp * -1 + ", " + cm.noncashexp * -1 + ", 0.00, 0.00, 0.00,0.00,0.00, 'Credit Memo Tool Write Off')";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "Update ARMatalloc set ARMLHLink = " + (lastLH + 1).ToString() + ",ARMFeeAdj = (([ARMFeeBld] - [ARMFeeRcvd] + [ARMFeeAdj]) * -1) + ARMFeeAdj ,ARMCshExpAdj = (([ARMCshExpBld] - [ARMCshExpRcvd] + [ARMCshExpAdj]) * -1) + ARMCshExpAdj " +
                  " ,ARMNCshExpAdj = (([ARMNCshExpBld] - [ARMNCshExpRcvd] + [ARMNCshExpAdj]) * -1) + ARMNCshExpAdj  ,[ARMBalDue] = 0 where ARMBillNbr = " + cm.inv + " and ARMMatter = " + cm.mat;

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "Update matter set MatAdjSinceLastBill = (MatAdjSinceLastBill + (" + cm.fees * -1 + "+" + cm.cashexp * -1 + "+ " + cm.noncashexp * -1 + ")) where matsysnbr = " + cm.mat;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            //MatAdjSinceLastBill = MatAdjSinceLastBill + adjustment


            SQL = "Update sysparam set spnbrvalue=spnbrvalue + 1 where spname='LastSysNbrLH'";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            int lastBatchNo = 0;

            sqlB = "select SpNbrValue from sysparam where spname = 'LastBatchCM'";
            DataSet spBatch1 = _jurisUtility.RecordsetFromSQL(sqlB);
            DataTable dtB1 = spBatch1.Tables[0];
            if (dtB1.Rows.Count == 0)
            { MessageBox.Show("Invalid sysparam data - LastBatchCM"); }
            else
            {
                foreach (DataRow dr in dtB1.Rows)
                {
                    lastBatchNo = Convert.ToInt32(dr["SpNbrValue"].ToString());
                }

            }

            string MYFolder = PYear + "-" + PNbr;

            SQL = "Insert into creditmemobatch( [CMBBatchNbr] ,[CMBComment] ,[CMBStatus] ,[CMBRecCount] ,[CMBEnteredBy] ,[CMBDateEntered] ,[CMBLastOpenedBy] ,[CMBLastOpenedDate] ,[CMBJEBatchNbr]) VALUES (" +
              lastBatchNo + 1 + ",'Write Off by JPS - Credit Memo Utility','U',1, (select empsysnbr from employee where EmpID = 'SMGR'), getdate(), " +
              " (select empsysnbr from employee where EmpID = 'SMGR'), getdate(), null)";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "Insert into creditmemo ( [CMBatchNbr] ,[CMRecNbr]  ,[CMLHLink] ,[CMBillNbr] ,[CMMatter] ,[CMComment] ,[CMDate] ,[CMPrdYear] ,[CMPrdNbr] ,[CMPreAdjFee] ,[CMFeeAdj] ,[CMPreAdjCshExp] ,[CMCshExpAdj] ,[CMPreAdjNCshExp] ,[CMNCshExpAdj],[CMPreAdjSurchg] ,[CMSurchgAdj] ,[CMPreAdjTax1] ,[CMTax1Adj] ,[CMPreAdjTax2] ,[CMTax2Adj] ,[CMPreAdjTax3] ,[CMTax3Adj] ,[CMPreAdjInterest] ,[CMInterestAdj],[CMPrintOption] ,[CMNarrative]) VALUES (" +
                                                   lastBatchNo + 1 + ",1," + (lastLH + 1).ToString() + "," + cm.inv + ", " + cm.mat + ", 'Write off by JPS - Credit Memo Tool', getdate(), " + PYear + ", " + PNbr + "," + cm.fees + "," + cm.fees * -1 + "," + cm.cashexp + "," + cm.cashexp * -1 + ", " + cm.noncashexp + "," + cm.noncashexp * -1 + ", 0.00, 0.00, 0.00,0.00,0.00,0.00, 0.00, 0.00, 0.00, 0.00, 'A', 'Credit Memo Tool Write Off')";

            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "Update arftaskalloc set ARFTAdj= ((arftactualamtbld - arftrcvd + arftadj ) * -1) + ARFTAdj" +
    " where arftmatter=" + cm.mat + " and arftbillnbr=" + cm.inv;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "select distinct ARFTTkpr from ARFTaskAlloc where ARFTBillNbr = " + cm.inv + " and ARFTMatter = " + cm.mat;

            DataSet dds = _jurisUtility.RecordsetFromSQL(SQL);

            foreach (DataRow rr in dds.Tables[0].Rows)
            {
                SQL = "Insert into CMFeeAlloc([CMFBatch]  ,[CMFRecNbr]  ,[CMFTkpr] ,[CMFBillNbr] ,[CMFMatter] ,[CMFPreAdj] ,[CMFAdj]) " +
                    " values (" + lastBatchNo + 1 + ", 1, " + rr[0].ToString() + ", " + cm.inv + ", " + cm.mat + ", (select sum(ARFTAdj) * -1 from ARFTaskAlloc where arftmatter=" + cm.mat + " and arftbillnbr=" + cm.inv + " and ARFTTkpr = " + rr[0].ToString() + ") , (select sum(ARFTAdj) from ARFTaskAlloc where arftmatter=" + cm.mat + " and arftbillnbr=" + cm.inv + " and ARFTTkpr = " + rr[0].ToString() + " ))";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }

            SQL = "Update arexpalloc set arepend= ((AREBldAmount - ARERcvd + AREAdj ) * -1) + arepend" +
" where AREMatter=" + cm.mat + " and AREBillNbr=" + cm.inv;
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "SELECT  [AREBillNbr] ,[AREMatter],[AREExpCd] ,[AREExpType] ,([AREBldAmount] - [ARERcvd] + [AREAdj]) as total FROM [ARExpAlloc] where AREBillNbr = " + cm.inv + " and AREMatter = " + cm.mat;

            dds.Clear();

            dds = _jurisUtility.RecordsetFromSQL(SQL);

            foreach (DataRow rr in dds.Tables[0].Rows)
            {

                SQL = "Insert into CMExpAlloc ([CMEBatch] ,[CMERecNbr] ,[CMEExpCd] ,[CMEExpType] ,[CMEBillNbr] ,[CMEMatter] ,[CMEPreAdj],[CMEAdj]) " +
                    " values (" + lastBatchNo + 1 + ", 1, '" + rr[2].ToString() + "', '" + rr[3].ToString() + "', " + rr[0].ToString() + ", " + rr[1].ToString() + ", " + rr[4].ToString() + ", " + (Convert.ToDouble(rr[4].ToString()) * -1).ToString()
                    + " )";
                _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            }


            SQL = "Update sysparam set spnbrvalue=spnbrvalue + 1 where spname='LastBatchCM'";
            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

            SQL = "select max(case when spname='CurAcctPrdYear' then cast(spnbrvalue as varchar(4)) else '' end) as PrdYear, " +
                               "max(Case when spname='CurAcctPrdNbr' then case when spnbrvalue<9 then '0' + cast(spnbrvalue as varchar(1)) else cast(spnbrvalue as varchar(2)) end  else '' end) as PrdNbr, " +
                               "max(case when spname='LastSysNbrDocTree' then spnbrvalue else 0 end) as DTree,max(case when spname='CfgMiscOpts' then substring(sptxtvalue,14,1) else 0 end) as DOrder from sysparam";
            DataSet myRSSysParm = _jurisUtility.RecordsetFromSQL(SQL);

            DataTable dtSP = myRSSysParm.Tables[0];

            if (dtSP.Rows.Count == 0)
            { MessageBox.Show("Incorrect SysParams"); }
            else
            {
                foreach (DataRow dr in dtSP.Rows)
                {
                    string LastSys = dr["DTree"].ToString();
                    DOrder = dr["DOrder"].ToString();
                    if (DOrder == "2")
                    {
                        string SPSql = "Select dtdocid from documenttree where dtparentid=37 and dtdocclass=5200 and dttitle='" + MYFolder + "'";
                        DataSet spMY = _jurisUtility.RecordsetFromSQL(SPSql);
                        DataTable dtMY = spMY.Tables[0];
                        if (dtMY.Rows.Count == 0)
                        {
                            string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                  "values((select max(dtdocid)  + 1 from documenttree), 'Y', 5200,'F', 37,'" + MYFolder + "') ";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'F', dtdocid,'SMGR'" +
                                " from documenttree where dtparentid=37 and dttitle='" + MYFolder + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'R', " +
                                " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37  and dttitle='" + MYFolder + "') and dttitle='SMGR')," +
                                "'JPS-Credit Memo Tool', " + lastBatchNo + 1;
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);


                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);
                        }
                        else
                        {
                            string SMGRSql = "Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37 and dttitle='" + MYFolder + "') and dttitle='SMGR'";
                            DataSet spSMGR = _jurisUtility.RecordsetFromSQL(SMGRSql);
                            DataTable dtSMGR = spSMGR.Tables[0];
                            if (dtSMGR.Rows.Count == 0)
                            {
                                string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                               "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'F', dtdocid,'SMGR'" +
                               " from documenttree where dtparentid=37 and dttitle='" + MYFolder + "'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                    "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'R', " +
                                    " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37 and dttitle='" + MYFolder + "')  and dttitle='SMGR')," +
                                    "'JPS-Credit Memo Tool', " + lastBatchNo + 1;
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            }
                            else
                            {
                                string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                    "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'R', " +
                                    " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37 and dttitle='" + MYFolder + "')  and dttitle='SMGR')," +
                                    "'JPS-Credit Memo Tool', " + lastBatchNo + 1;
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            }
                        }
                    }
                    else
                    {
                        string SPSql = "Select dtdocid from documenttree where dtparentid=37 and dtdocclass=5200 and dttitle='SMGR'";
                        DataSet spMY = _jurisUtility.RecordsetFromSQL(SPSql);
                        DataTable dtMY = spMY.Tables[0];
                        if (dtMY.Rows.Count == 0)
                        {
                            string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                  "values ((select max(dtdocid)  + 1 from documenttree), 'Y', 5200,'F', 37,'SMGR') ";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                                "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'F', dtdocid,'" + MYFolder + "'" +
                                " from documenttree where dtparentid=37 and dttitle='SMGR'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);

                            s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'R', " +
                                " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37  and dttitle='SMGR') and dttitle='" + MYFolder + "')," +
                                "'JPS-Credit Memo Tool', " + lastBatchNo + 1;
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);


                            s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s2);
                        }
                        else
                        {
                            string SMGRSql = "Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37  and dttitle='SMGR') and dttitle='" + MYFolder + "'";
                            DataSet spSMGR = _jurisUtility.RecordsetFromSQL(SMGRSql);
                            DataTable dtSMGR = spSMGR.Tables[0];
                            if (dtSMGR.Rows.Count == 0)
                            {
                                string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle) " +
                               "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'F', dtdocid,'" + MYFolder + "'" +
                               " from documenttree where dtparentid=37 and dttitle='SMGR'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);

                                s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                    "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'R', " +
                                    " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37   and dttitle='SMGR')and dttitle='" + MYFolder + "')," +
                                    "'JPS-Credit Memo Tool', " + lastBatchNo + 1;
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            }
                            else
                            {
                                string s2 = "Insert into documenttree(dtdocid, dtsystemcreated, dtdocclass, dtdoctype, dtparentid, dttitle, dtkeyL) " +
                                    "select (select max(dtdocid) from documenttree) + 1, 'Y', 5200,'R', " +
                                    " (Select dtdocid from documenttree where dtparentid=(Select dtdocid from documenttree where dtparentid=37 and dttitle='SMGR') and dttitle='" + MYFolder + "') ," +
                                    "'JPS-Credit Memo Tool', " + lastBatchNo + 1;
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);


                                s2 = "Update sysparam set spnbrvalue=(select max(dtdocid) from documenttree) where spname='LastSysNbrDocTree'";
                                _jurisUtility.ExecuteNonQueryCommand(0, s2);
                            }
                        }

                    }
                }
            }

            return lastBatchNo + 1;


        }




        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion





        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            string invs = "";
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Browse for Invoice File";
            dlg.DefaultExt = "txt";
            dlg.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            dlg.Multiselect = false;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fileName = dlg.FileName;
                

                String line;
                try
                {
                    //Pass the file path and file name to the StreamReader constructor
                    StreamReader sr = new StreamReader(fileName);

                    //Read the first line of text
                    line = sr.ReadLine();

                    //Continue to read until you reach end of file
                    while (line != null)
                    {
                        //write the lie to console window
                        invoices.Add(line);
                        //Read the next line
                        line = sr.ReadLine();
                    }

                    //close the file
                    sr.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                }

                if (invoices != null && invoices.Count > 0)
                {
                    
                    foreach (string inv in invoices)
                    {
                        invs = invs + inv + ",";
                    }
                    invs = invs.TrimEnd(',');
                    string sqlB = "select ARMBillNbr as BillNumber, dbo.jfn_FormatClientCode(clicode) as ClientCode, dbo.jfn_FormatMatterCode(matcode) as MatterCode, cast(sum([ARMFeeBld] - [ARMFeeRcvd] + [ARMFeeAdj]) as decimal(15,2)) as Fees, " +
                    " cast(sum([ARMCshExpBld] - [ARMCshExpRcvd] + [ARMCshExpAdj]) as decimal(15, 2)) as CashExp,  cast(sum([ARMNCshExpBld] - [ARMNCshExpRcvd] + [ARMNCshExpAdj]) as decimal(15,2)) as NonCashExp, " +
                    " cast(sum(([ARMCshExpBld] - [ARMCshExpRcvd] + [ARMCshExpAdj]) + ([ARMNCshExpBld] - [ARMNCshExpRcvd] + [ARMNCshExpAdj]) + ([ARMFeeBld] - [ARMFeeRcvd] + [ARMFeeAdj])) as decimal(15,2)) as Total, matsysnbr as matID, " +
                    " cast(sum([ARMFeeAdj]) as decimal(15,2)) as FeeAdj, cast(sum([ARMCshExpAdj]) as decimal(15, 2)) as CashExpAdj, cast(sum([ARMNCshExpAdj]) as decimal(15,2)) as NonCashExpAdj " +
                    " from ARMatAlloc  " +
                    " inner join matter on matsysnbr = ARMMatter " +
                    " inner join client on clisysnbr = matclinbr " +
                    " where ARMBillNbr in (" + invs + ") " +
                    " group by ARMBillNbr, clicode,matcode, matsysnbr " + 
                    " having sum(([ARMCshExpBld] - [ARMCshExpRcvd] + [ARMCshExpAdj]) + ([ARMNCshExpBld] - [ARMNCshExpRcvd] + [ARMNCshExpAdj]) + ([ARMFeeBld] - [ARMFeeRcvd] + [ARMFeeAdj])) <> 0";
                    DataSet spBatch = _jurisUtility.RecordsetFromSQL(sqlB);
                    if (spBatch.Tables[0].Rows.Count == 0)
                    { MessageBox.Show("There were no records with those invoice numbers", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                    else
                    {
                        DataTable dtSP = spBatch.Tables[0];

                        if (dtSP.Rows.Count == 0)
                        { MessageBox.Show("Incorrect SysParams"); }
                        else
                        {
                            CreditMemo mm = null;
                            foreach (DataRow dr in dtSP.Rows)
                            {
                                mm = new CreditMemo();
                                mm.inv = Convert.ToInt32(dr["BillNumber"].ToString());
                                mm.mat = Convert.ToInt32(dr["matID"].ToString());
                                mm.cashexp = Convert.ToDouble(dr["CashExp"].ToString());
                                mm.fees = Convert.ToDouble(dr["Fees"].ToString());;
                                mm.noncashexp = Convert.ToDouble(dr["NonCashExp"].ToString());
                                mm.FeeAdj = Convert.ToDouble(dr["FeeAdj"].ToString());
                                mm.CashExpAdj = Convert.ToDouble(dr["CashExpAdj"].ToString());
                                mm.NonCashExpAdj = Convert.ToDouble(dr["NonCashExpAdj"].ToString());
                                memos.Add(mm);
                            }
                        }
                        dataGridView1.DataSource = spBatch.Tables[0];
                    }

                }
                else
                {
                    MessageBox.Show("There were no valid invoice numbers in the selected text" + "\r\n" + " file or none of the invoices have a balance", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }



            }

 }


    }
}
