using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Drawing;
using Newtonsoft.Json;
using CarlosAg.ExcelXmlWriter;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using System.Linq.Expressions;
using AgencyPartnerPortalDB.DatabaseLayer;

namespace AgencyPartnerPortal_v1._0.CommonFiles
{
    public partial class rptgroups_downloadfile : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string idparms = Session["Rpt1params"].ToString();
            //string idparms = Request.QueryString["idparms"].ToString();
            List<params1> parmsData = WM_DecryptUrl(idparms);

            if (parmsData[0].FrmScreen == "AdmnReport1")
            {

                if (parmsData[0].RptTyp == "Excel")
                {
                    GenerateExcel(parmsData[0].ReqIDArry, parmsData[0].Rpt1params);
                }
                else if (parmsData[0].RptTyp == "Pdf")
                {
                }
            }
        }

        void GenerateExcel(List<ReqParams> ReqIDArry, List<Rpt1params> oRpt1params)
        {

            try
            {
                #region DEMOGRAPHICS MASTER TABLES
                /********** DEMOGRAPHICS MASTERS********************/

                DataSet _dsMasters = Commonfunctions.BuildDemographicMasters();

                DataTable dtGenderData = _dsMasters.Tables["Gender"];
                DataTable dtRaceData = _dsMasters.Tables["Race"];
                DataTable dtEthnicityData = _dsMasters.Tables["Ethnicity"];
                DataTable dtCountyData = _dsMasters.Tables["County"];
                DataTable dtCityData = _dsMasters.Tables["City"];
                DataTable dtLanguagesData = _dsMasters.Tables["Languages"];

                /***************************************************/
                #endregion

                DataTable dtgrpReq = new DataTable();
                DataTable dtQuestions = new DataTable();

                DataSet dsServReqstData = AgencyPartnerPortalDB.DatabaseLayer.Capsystems.GetReport1(oRpt1params); // Main Service Request Data
                dtQuestions = dsServReqstData.Tables[1].Copy();
                DataSet dsReviewData = AgencyPartnerPortalDB.DatabaseLayer.Capsystems.APRREVIEW_GET("", "BYID"); // Main All Request Reviewd Data
                //******************* ***** ********* ***** * ***************************//

                Workbook book = new Workbook();

                #region Excel Cell Styles

                /********************** SET EXCEL FONT PROPERTIES ******************************/

                CarlosAg_Excel_Properties oCarlosAg_Excel_Properties = new CarlosAg_Excel_Properties();

                oCarlosAg_Excel_Properties.sxlbook = book;
                oCarlosAg_Excel_Properties.sxlTitleFont = "Century Gothic";
                oCarlosAg_Excel_Properties.sxlbodyFont = "calibri";
                /*******************************************************************************/
                oCarlosAg_Excel_Properties.getCarlosAg_Excel_Properties();

                string xAgyHeaderCellstyle = oCarlosAg_Excel_Properties.xfnCELL_STYLE(book, "xAgyHeaderCellstyle", oCarlosAg_Excel_Properties.sxlbodyFont, 8, "#FFFFFF", true, "#305496", "left", 1, 1, 1, 1, "#305496").ToString();
                string gxlBCC_sp_cr = oCarlosAg_Excel_Properties.xfnCELL_STYLE(book, "gxlBCC_sp_cr", oCarlosAg_Excel_Properties.sxlbodyFont, 8, "#000000", true, "#d3e6f5", "center", 1, 1, 1, 1, "#d3e6f5").ToString();
                string gxlGCC_sp_cr = oCarlosAg_Excel_Properties.xfnCELL_STYLE(book, "gxlGCC_sp_cr", oCarlosAg_Excel_Properties.sxlbodyFont, 8, "#000000", true, "#bee396", "center", 1, 1, 1, 1, "#bee396").ToString();

                #endregion

                Worksheet sheet;
                WorksheetRow excelrow;
                WorksheetCell cell;
                /*************************************************/

                int Mergcells = 9;


                foreach (ReqParams ReqData in ReqIDArry)
                {
                    string SPID = ReqData.SPID.ToString();

                    string SPgrpType = ReqData.grpType.ToString();
                    string SPName = ReqData.ServName.ToString();
                    string _strTitle = SPName;



                    string _strProgram = SPgrpType + " - " + SPName;

                    string _strSheetName = "";
                    Regex reg = new Regex("[*'\"/,_&#^@]");
                    if (_strProgram.Length > 24)
                    {
                        _strSheetName = _strProgram.Substring(0, 24);
                        _strSheetName = reg.Replace(_strSheetName, string.Empty);

                        //_strSheetName = _strSheetName + " - " + (Convert.ToInt32(_strFisicalYear.Substring(2, 2)) + 1).ToString();
                    }
                    else
                    {
                        _strProgram = reg.Replace(_strProgram, string.Empty);
                        _strSheetName = _strProgram;
                    }

                    sheet = book.Worksheets.Add(_strSheetName);
                    sheet.Options.GridLineColor = "#FFFFFF";

                    dtgrpReq = new DataTable();
                    // DataSet ds = AgencyPartnerPortalDB.DatabaseLayer.Capsystems.GetSPMeasures("", SPID, "SEREQBYID");
                    DataTable _dt = new DataTable();
                    DataRow[] drservReqdata = dsServReqstData.Tables[0].Select("APR_SEREQ_SP='" + SPID + "' AND APR_SEREQ_SP_TYPE='" + SPgrpType + "'");
                    if (drservReqdata.Length > 0)
                        _dt = drservReqdata.CopyToDataTable();

                    DataTable dtAll = BuildDataSetWithAllTables(_dt, dtQuestions, _dsMasters, SPID, SPgrpType);
                    DataTable dtAgeData = CommonFiles.Commonfunctions.BuildAgeMaster(SPID);
                    DataSet dsQueMast = AgencyPartnerPortalDB.DatabaseLayer.Capsystems.CAPS_AGCYSERQASSOC_GET(SPID, "", "SERQASSOC");
                    if (dtAll.Rows.Count > 0)
                    {
                        DataRow[] drTemp = dtAll.Select("APR_SEREQ_SP_TYPE='" + SPgrpType + "'");
                        if (drTemp.Length > 0)
                        {
                            dtgrpReq = drTemp.CopyToDataTable();
                            //dtQuestions = ds.Tables[1].Copy();
                        }
                    }

                    int genderCnt = 0;
                    int ageCnt = 0;
                    int raceCnt = 0;
                    int ethinicityCnt = 0;
                    int countyCnt = 0;
                    int cityCnt = 0;
                    int languageCnt = 0;

                    int questionsCnt = 0;
                    int serviceCnt = 0;
                    DataTable dtServices = new DataTable();
                    dtServices = JsonConvert.DeserializeObject<DataTable>(servicerequest.WM_GetCaseSP2(SPID));

                    if (dtgrpReq.Rows.Count > 0)
                    {
                        #region Groups Report Code
                        if (SPgrpType == "G")
                        {

                            if (SPID == "19")
                                serviceCnt = 7;
                            else
                                serviceCnt = dtServices.Rows.Count;



                            genderCnt = dtGenderData.Rows.Count;
                            ageCnt = dtAgeData.Rows.Count;
                            raceCnt = dtRaceData.Rows.Count;
                            ethinicityCnt = dtEthnicityData.Rows.Count;
                            countyCnt = dtCountyData.Rows.Count;
                            cityCnt = dtCityData.Rows.Count;
                            languageCnt = dtLanguagesData.Rows.Count;

                            questionsCnt = dsQueMast.Tables[0].Rows.Count;

                            #region ADD COLUMNS

                            sheet.Table.Columns.Add(new WorksheetColumn(170));  //Agency Column
                            sheet.Table.Columns.Add(new WorksheetColumn(40));   // Request No
                            sheet.Table.Columns.Add(new WorksheetColumn(60));   //Request Date
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   //Completion Date
                            sheet.Table.Columns.Add(new WorksheetColumn(60));   //Change Date
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   //Change Operator

                            //** Request Columns **//
                            for (int x = 0; x < serviceCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(60));
                            }

                            sheet.Table.Columns.Add(new WorksheetColumn(50));
                            /*******************************************************************/

                            //** Actual Distribution Columns **//
                            if (SPID != "19")
                            {
                                for (int x = 0; x < serviceCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(60));
                                }
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Qty
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Transaction Units
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Unit Price
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Value


                                sheet.Table.Columns.Add(new WorksheetColumn(50));   // Description
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Individuals Served
                                sheet.Table.Columns.Add(new WorksheetColumn(70));   //Demographics Update Required
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Status
                                /***********************************************************************/
                            }
                            else
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(60));   //Item
                                sheet.Table.Columns.Add(new WorksheetColumn(60));   //Service

                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Qty
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Transaction Units
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Unit Price
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Value


                                sheet.Table.Columns.Add(new WorksheetColumn(50));   // Description
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Status
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Individuals Served
                                /***********************************************************************/
                            }

                            /** Representative Details**/
                            sheet.Table.Columns.Add(new WorksheetColumn(70));
                            sheet.Table.Columns.Add(new WorksheetColumn(70));
                            sheet.Table.Columns.Add(new WorksheetColumn(70));
                            sheet.Table.Columns.Add(new WorksheetColumn(70));
                            sheet.Table.Columns.Add(new WorksheetColumn(140));
                            sheet.Table.Columns.Add(new WorksheetColumn(120));

                            sheet.Table.Columns.Add(new WorksheetColumn(60));

                            /** Gender Details**/
                            for (int x = 0; x < genderCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(50));
                            }
                            /** Age Details**/
                            for (int x = 0; x < ageCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(40));
                            }
                            /** Race Details**/
                            for (int x = 0; x < raceCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(110));
                            }
                            /** Ethnicity Details**/
                            for (int x = 0; x < ethinicityCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(100));
                            }
                            /** County Details**/
                            for (int x = 0; x < countyCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(70));
                            }
                            /** City Details**/
                            for (int x = 0; x < cityCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(70));
                            }


                            /*Questions Detaisl*/

                            for (int x = 0; x < questionsCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(220));
                            }

                            #endregion


                            #region TITLE BLOCK 1
                            //string _strTitle = dtgrpReq.Rows[0]["OSERVICES"].ToString();
                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add("Agency Partners Reports - Groups ", DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle1);
                            cell.MergeAcross = Mergcells;

                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add(_strTitle, DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle2);
                            cell.MergeAcross = Mergcells;

                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlEMPTC);
                            cell.MergeAcross = Mergcells;

                            #endregion

                            #region COLUMNS BLOCK 2
                            excelrow = sheet.Table.Rows.Add();
                            // cell = excelrow.Cells.Add("slno", DataType.String, xAgyHeaderCellstyle);
                            cell = excelrow.Cells.Add("Agency", DataType.String, xAgyHeaderCellstyle);
                            //  cell.MergeDown = 2;
                            cell = excelrow.Cells.Add("Request No.", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Request Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Completion Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Change Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Change Operator", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);

                            if (serviceCnt > 0)
                            {
                                cell = excelrow.Cells.Add("REQUEST", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC_sp);
                                cell.MergeAcross = serviceCnt;
                            }

                            if (serviceCnt > 0)
                            {
                                cell = excelrow.Cells.Add("ACTUAL DISTRIBUTION", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC_sp);
                                if (SPID != "19")
                                    cell.MergeAcross = serviceCnt + 7;
                                else
                                    cell.MergeAcross = 2 + 6;
                            }

                            cell = excelrow.Cells.Add("REPRESENTATIVE DETAILS", DataType.String, oCarlosAg_Excel_Properties.gxlBRCHC_sp);
                            cell.MergeAcross = 5;

                            cell = excelrow.Cells.Add("Target Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);


                            cell = excelrow.Cells.Add("GENDER", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = genderCnt - 1;

                            cell = excelrow.Cells.Add("AGE%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = ageCnt - 1;

                            cell = excelrow.Cells.Add("RACE%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = raceCnt - 1;

                            cell = excelrow.Cells.Add("ETHNICITY%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = ethinicityCnt - 1;

                            cell = excelrow.Cells.Add("COUNTY%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = countyCnt - 1;

                            cell = excelrow.Cells.Add("CITY%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = cityCnt - 1;

                            cell = excelrow.Cells.Add(_strTitle + " " + "QUESTIONS", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell.MergeAcross = questionsCnt - 1;

                            #endregion

                            #region COLUMNS-2 BLOCK 3
                            excelrow = sheet.Table.Rows.Add();
                            // cell = excelrow.Cells.Add("", DataType.String, xAgyHeaderCellstyle);
                            cell = excelrow.Cells.Add("", DataType.String, xAgyHeaderCellstyle);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);

                            //REQUEST
                            if (serviceCnt > 0)
                            {

                                if (SPID != "19")
                                {
                                    foreach (DataRow drserHeads in dtServices.Rows)
                                    {
                                        cell = excelrow.Cells.Add(drserHeads["CAMS_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    }
                                }
                                else
                                {
                                    cell = excelrow.Cells.Add("Season", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Demographic Group", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Clothing Type/Item Descp.".ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Size", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Total", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Units", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Additional Info", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                }
                                cell = excelrow.Cells.Add("Total Ind. Served", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                            }

                            //ACTUAL DISTRIBUTION
                            if (serviceCnt > 0)
                            {
                                if (SPID != "19")
                                {
                                    foreach (DataRow drserHeads in dtServices.Rows)
                                    {
                                        cell = excelrow.Cells.Add(drserHeads["CAMS_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    }
                                }
                                else
                                {
                                    cell = excelrow.Cells.Add("Item", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Service", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                }

                                cell = excelrow.Cells.Add("Total Qty", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                cell = excelrow.Cells.Add("Transaction Units", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                cell = excelrow.Cells.Add("Unit Price", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                cell = excelrow.Cells.Add("Total Value", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                cell = excelrow.Cells.Add("Description", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);

                                if (SPID != "19")
                                {
                                    cell = excelrow.Cells.Add("Total Ind. Served", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Demographics Update Require", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Status", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                }
                                else
                                {
                                    cell = excelrow.Cells.Add("Status", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Total Ind. Served", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                }


                            }
                            //REPRESENTATIVE DETAILS
                            cell = excelrow.Cells.Add("Name", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Last Name", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Phone", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Phone2", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Email", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Position", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);

                            //Target Date
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNLHC);


                            //GENDER DISTRIBUTION
                            foreach (DataRow drRowsHead in dtGenderData.Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            }

                            //AGE DISTRIBUTION
                            foreach (DataRow drRowsHead in dtAgeData.Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["id"].ToString() + " yrs.", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            }

                            //Race DISTRIBUTION
                            foreach (DataRow drRowsHead in dtRaceData.Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            }

                            //Ethnicity DISTRIBUTION
                            foreach (DataRow drRowsHead in dtEthnicityData.Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            }

                            //County DISTRIBUTION
                            foreach (DataRow drRowsHead in dtCountyData.Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            }
                            //City DISTRIBUTION
                            foreach (DataRow drRowsHead in dtCityData.Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["SQR_RESP_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            }


                            //QUESTIONS 
                            foreach (DataRow drRowsHead in dsQueMast.Tables[0].Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["AGYQ_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNLHC);
                            }

                            #endregion

                            #region ROWS DATA BLOCK4

                            string strAgency = "";
                            string ReqDate = "";
                            string ReqCompDate = "";
                            string strChangeOperator = "";
                            string RequestID = "";
                            string ReqType = "";
                            string ReqSPID = "";
                            string ReqStatus = "";

                            int RowID = 5; int TempRowCount = 0;
                            int chkIsfirst = 0;
                            foreach (DataRow drserDets in dtgrpReq.Rows)
                            {
                                if (chkIsfirst > -1)
                                {
                                    strAgency = drserDets["AGYNAME"].ToString();
                                    ReqDate = Convert.ToDateTime(drserDets["APR_SEREQ_DATE_ADD"]).ToString("MM/dd/yyyy");
                                    ReqCompDate = Convert.ToDateTime(drserDets["APR_SEREQ_DATE_LSTC"]).ToString("MM/dd/yyyy");
                                    strChangeOperator = drserDets["APR_SEREQ_LSTC_OPERATOR"].ToString();
                                    RequestID = drserDets["APR_SEREQ_ID"].ToString();
                                    ReqType = drserDets["APR_SEREQ_SP_TYPE"].ToString();
                                    ReqSPID = drserDets["APR_SEREQ_SP"].ToString();
                                    ReqStatus = drserDets["SERSTATUS"].ToString();

                                    DataTable _dt19Services = JsonConvert.DeserializeObject<DataSet>(drserDets["APR_SERQ_DEMOGRAPH"].ToString()).Tables[0];
                                    int _19SPMergeCount = _dt19Services.Rows.Count - 1;


                                    excelrow = sheet.Table.Rows.Add();
                                    // excelrow.Index = RowID;
                                    // cell = excelrow.Cells.Add(RequestID, DataType.String, oCarlosAg_Excel_Properties.xNR_Left_Cellstyle);
                                    cell = excelrow.Cells.Add(strAgency, DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                    if (SPID == "19")
                                        cell.MergeDown = _19SPMergeCount;

                                    cell = excelrow.Cells.Add(RequestID, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    if (SPID == "19")
                                        cell.MergeDown = _19SPMergeCount;

                                    cell = excelrow.Cells.Add(ReqDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    if (SPID == "19")
                                        cell.MergeDown = _19SPMergeCount;

                                    cell = excelrow.Cells.Add(ReqCompDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    if (SPID == "19")
                                        cell.MergeDown = _19SPMergeCount;

                                    cell = excelrow.Cells.Add(ReqCompDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    if (SPID == "19")
                                        cell.MergeDown = _19SPMergeCount;

                                    cell = excelrow.Cells.Add(strChangeOperator, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    if (SPID == "19")
                                        cell.MergeDown = _19SPMergeCount;

                                    ///********************************** --- REQUEST    ---    **************************************///
                                    ///
                                    string str19IndvCount = "";
                                    if (SPID != "19")
                                    {
                                        if (serviceCnt > 0)
                                        {
                                            string[] Servarray = drserDets["APR_SERQ_SERVICES"].ToString().Split(',');
                                            string[] Countarray = drserDets["APR_SERQ_INDIV_CNT"].ToString().Split(',');
                                            if (Servarray.Length > 0)
                                            {
                                                int ttlIndserv = 0;
                                                foreach (DataRow drserHeads in dtServices.Rows)
                                                {
                                                    string strServCode = drserHeads["SP2_CAMS_CODE"].ToString().Trim();
                                                    var serIndex = Array.IndexOf(Servarray, strServCode);
                                                    if (serIndex > -1)
                                                    {
                                                        string serValue = Countarray[serIndex];
                                                        cell = excelrow.Cells.Add(serValue.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                        ttlIndserv += Convert.ToInt32(serValue);
                                                    }
                                                    else
                                                    {
                                                        cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    }
                                                }
                                                cell = excelrow.Cells.Add(ttlIndserv.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC_cr);
                                            }
                                            else
                                            {
                                                foreach (DataRow drserHeads in dtServices.Rows)
                                                {
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                }
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC_cr);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (_dt19Services.Rows.Count > 0)
                                        {

                                            int _MiddleRowNo = MiddleRowNumb(_dt19Services.Rows.Count);
                                            int t = 0;
                                            foreach (DataRow dr19Services in _dt19Services.Rows)
                                            {
                                                if (t > 0)
                                                {
                                                    excelrow = sheet.Table.Rows.Add();
                                                }

                                                cell = excelrow.Cells.Add(dr19Services["SeasonReq"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                if (t > 0)
                                                    cell.Index = 7;
                                                cell = excelrow.Cells.Add(dr19Services["DemoGrp"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr19Services["ItemDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr19Services["Size"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr19Services["Total"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr19Services["Units"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr19Services["Additionalinfo"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);

                                                if (t == (_MiddleRowNo))
                                                {
                                                    cell = excelrow.Cells.Add(drserDets["APR_SERQ_INDIV_CNT"].ToString(), DataType.String, gxlBCC_sp_cr);
                                                    //cell.MergeDown = _19SPMergeCount;
                                                    str19IndvCount = drserDets["APR_SERQ_INDIV_CNT"].ToString();
                                                }
                                                else
                                                    cell = excelrow.Cells.Add("", DataType.String, gxlBCC_sp_cr);
                                                t++;
                                            }


                                        }

                                    }

                                    ///*****************************************************************************************************************///

                                    ///********************************** --- ACTUAL DISTRIBUTION    ---    **************************************///

                                    DataRow[] drReviewDets = dsReviewData.Tables[0].Select("APR_REV_REQ_ID=" + RequestID + "");
                                    if (SPID != "19")
                                    {
                                        if (drReviewDets.Length > 0)
                                        {
                                            DataTable dtReviewDets = drReviewDets.CopyToDataTable();
                                            decimal ttldistributon = 0;

                                            decimal ttlQTY = 0;
                                            foreach (DataRow drserHeads in dtServices.Rows)
                                            {
                                                string strServCode = drserHeads["SP2_CAMS_CODE"].ToString().Trim();
                                                DataRow[] drRevArry = dtReviewDets.Select("APR_REV_SERVICE='" + strServCode + "'");
                                                if (drRevArry.Length > 0)
                                                {
                                                    foreach (DataRow drRDets in drRevArry)
                                                    {
                                                        string Qty = drRDets["APR_REV_QTY"].ToString().Trim();
                                                        cell = excelrow.Cells.Add(Qty.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                        ttlQTY += Qty == "" ? 0 : Convert.ToDecimal(Qty);
                                                    }

                                                }
                                                else
                                                {
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                }
                                            }
                                            cell = excelrow.Cells.Add(ttlQTY.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);


                                            cell = excelrow.Cells.Add("FAB Bundle", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("50", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            decimal ttlValue = 50 * ttlQTY;
                                            cell = excelrow.Cells.Add(ttlValue.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                            cell = excelrow.Cells.Add(ttlQTY.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC_cr);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(ReqStatus.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                        }
                                        else
                                        {
                                            foreach (DataRow drserHeads in dtServices.Rows)
                                            {
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            }
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC_cr);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(ReqStatus.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                        }
                                    }
                                    else
                                    {
                                        if (drReviewDets.Length > 0)
                                        {
                                            DataTable _dt19RevewData = drReviewDets.CopyToDataTable();
                                            int _MiddleRowNo = MiddleRowNumb(_dt19Services.Rows.Count);
                                            int i = 0;
                                            for (int x = 0; x < (_dt19Services.Rows.Count); x++)
                                            {
                                                int _cRowID = (RowID + i);

                                                string _fItemDesc = _dt19Services.Rows[x]["ItemDesc"].ToString() == "-" ? "" : (_dt19Services.Rows[x]["ItemDesc"].ToString() + "|");
                                                string _fSeason = _dt19Services.Rows[x]["SeasonReq"].ToString() == "-" ? "" : (_dt19Services.Rows[x]["SeasonReq"].ToString() + "|");
                                                string _fDemoGrp = _dt19Services.Rows[x]["Demogrp"].ToString() == "-" ? "" : (_dt19Services.Rows[x]["Demogrp"].ToString() + "|");
                                                string _fSize = _dt19Services.Rows[x]["Size"].ToString() == "-" ? "" : (_dt19Services.Rows[x]["Size"].ToString() + "|");

                                                string SearchString = _fItemDesc + _fSeason + _fDemoGrp + _fSize;

                                                DataRow[] _dr19RevData = _dt19RevewData.Select("APR_REV_ITEM_DESC='" + SearchString.TrimEnd('|') + "'");

                                                if (_dr19RevData.Length > 0)
                                                {
                                                    string strServDesc = "";
                                                    DataRow[] drstrServDesc = dtServices.Select("SP2_CAMS_CODE='" + _dr19RevData[0]["APR_REV_SERVICE"].ToString() + "'");
                                                    if (drstrServDesc.Length > 0)
                                                    {
                                                        strServDesc = drstrServDesc[0]["CAMS_DESC"].ToString();
                                                    }

                                                    //var strServDesc = (from r in dtServices.AsEnumerable()
                                                    //                     where r.Field<string>("SP2_CAMS_CODE") == _dr19RevData[0]["APR_REV_SERVICE"].ToString()
                                                    //                     select r.Field<string>("CAMS_DESC")).First<string>();

                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["APR_REV_ITEM_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(strServDesc, DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["APR_REV_QTY"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["APR_REV_TRANS_UNIT"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["APR_REV_UNITPRICE"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["APR_REV_TOTAL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["APR_REV_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr19RevData[0]["SERSTATUS"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    if (i == (_MiddleRowNo))
                                                    {
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(str19IndvCount, DataType.String, gxlGCC_sp_cr);
                                                    }
                                                    else
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, gxlGCC_sp_cr);
                                                }
                                                else
                                                {
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                                    if (i == (_MiddleRowNo))
                                                    {
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(str19IndvCount, DataType.String, gxlGCC_sp_cr);
                                                    }
                                                    else
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, gxlGCC_sp_cr);
                                                }

                                                i++;
                                            }
                                        }
                                        else
                                        {
                                            int i = 0;
                                            for (int x = 0; x < (_dt19Services.Rows.Count); x++)
                                            {

                                                int _cRowID = (RowID + i);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, gxlGCC_sp_cr);

                                                i++;
                                            }
                                        }
                                    }
                                    ///*****************************************************************************************************************///

                                    ///****************** SET FIXED EXCEL ROW ************************////
                                    if (SPID == "19")
                                        excelrow = sheet.Table.Rows[RowID];
                                    ////******************************************////

                                    ///********************************************* Representative Details ************************************************///
                                    ///

                                    if (SPID != "19")
                                    {
                                        cell = excelrow.Cells.Add(drserDets["REPFNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["REPLNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["PHONE1"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["PHONE2"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["REPEMAIL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["REPPOST"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                    }
                                    else
                                    {
                                        cell = excelrow.Cells.Add(drserDets["REPFNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _19SPMergeCount;
                                        cell = excelrow.Cells.Add(drserDets["REPLNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _19SPMergeCount;
                                        cell = excelrow.Cells.Add(drserDets["PHONE1"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _19SPMergeCount;
                                        cell = excelrow.Cells.Add(drserDets["PHONE2"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _19SPMergeCount;
                                        cell = excelrow.Cells.Add(drserDets["REPEMAIL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _19SPMergeCount;
                                        cell = excelrow.Cells.Add(drserDets["REPPOST"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _19SPMergeCount;
                                    }

                                    ///*****************************************************************************************************************///
                                    ///
                                    if (SPID != "19")
                                        cell = excelrow.Cells.Add(Convert.ToDateTime(drserDets["APR_SEREQ_TARGET_DT"].ToString()).ToString("MM/dd/yyyy"), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    else
                                    {
                                        cell = excelrow.Cells.Add(Convert.ToDateTime(drserDets["APR_SEREQ_TARGET_DT"].ToString()).ToString("MM/dd/yyyy"), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _19SPMergeCount;
                                    }



                                    ////GENDER DISTRIBUTION
                                    foreach (DataRow drRowsHead in dtGenderData.Rows)
                                    {
                                        string code = "gen_" + drRowsHead["Code"].ToString();
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }

                                    //AGE DISTRIBUTION
                                    foreach (DataRow drRowsHead in dtAgeData.Rows)
                                    {
                                        string code = "age_" + drRowsHead["id"].ToString();
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }

                                    //Race DISTRIBUTION
                                    foreach (DataRow drRowsHead in dtRaceData.Rows)
                                    {
                                        string code = "rac_" + drRowsHead["code"].ToString();
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }

                                    //Ethnicity DISTRIBUTION
                                    foreach (DataRow drRowsHead in dtEthnicityData.Rows)
                                    {
                                        string code = "eth_" + drRowsHead["code"].ToString();
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }

                                    //County DISTRIBUTION
                                    foreach (DataRow drRowsHead in dtCountyData.Rows)
                                    {
                                        string code = "cou_" + drRowsHead["code"].ToString();
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }
                                    //City DISTRIBUTION
                                    foreach (DataRow drRowsHead in dtCityData.Rows)
                                    {
                                        string code = "cit_" + drRowsHead["SQR_RESP_CODE"].ToString().Trim().Replace(" ", "_");
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }


                                    // QUESTIONS DETAILS
                                    foreach (DataRow drRowsQ in dsQueMast.Tables[0].Rows)
                                    {
                                        string code = "que_" + drRowsQ["AGYQ_CODE"].ToString();
                                        if (SPID != "19")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                        else
                                        {
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                            cell.MergeDown = _19SPMergeCount;
                                        }
                                    }


                                    excelrow = sheet.Table.Rows.Add();
                                    excelrow.Height = 5;
                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlEMPTC);
                                    cell.MergeAcross = Mergcells;

                                    ///**************************  SET ROWID *********************************////

                                    if (chkIsfirst == 0)
                                    {
                                        TempRowCount = RowID + _19SPMergeCount + 1 + 1;
                                    }
                                    else if (chkIsfirst > 0)
                                    {

                                        TempRowCount = (TempRowCount + _19SPMergeCount + 1) + 1;
                                    }



                                    //**************************************************************************///
                                }
                                RowID = (TempRowCount == 0 ? 5 : TempRowCount);
                                chkIsfirst++;
                            }

                            #endregion
                        }
                        #endregion
                        #region Individuals Report Code
                        else
                        {
                            DataTable dt23DemographData = new DataTable();
                            serviceCnt = dtServices.Rows.Count;
                            if (SPID != "17")
                            {
                                genderCnt = dtGenderData.Rows.Count;
                                ageCnt = dtAgeData.Rows.Count;
                                raceCnt = dtRaceData.Rows.Count;
                                ethinicityCnt = dtEthnicityData.Rows.Count;
                                countyCnt = dtCountyData.Rows.Count;
                                cityCnt = dtCityData.Rows.Count;
                                languageCnt = dtLanguagesData.Rows.Count;
                            }
                            questionsCnt = dsQueMast.Tables[0].Rows.Count;

                            if (SPID == "23")
                            {
                                dt23DemographData = BuildDataSet23WithAllTables(_dt, dtQuestions, _dsMasters, SPID, SPgrpType);
                            }

                            #region ADD COLUMNS

                            sheet.Table.Columns.Add(new WorksheetColumn(170));  //Agency Column
                            sheet.Table.Columns.Add(new WorksheetColumn(40));   // Request No
                            sheet.Table.Columns.Add(new WorksheetColumn(60));   //Request Date
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   //Completion Date
                            sheet.Table.Columns.Add(new WorksheetColumn(60));   //Change Date
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   //Change Operator

                            //** Request Columns **//
                            if (SPID == "17")
                            {
                                serviceCnt = serviceCnt + 5;
                            }
                            else if (SPID == "23")
                            {
                                serviceCnt = 12;
                            }
                            for (int x = 0; x < serviceCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(60));
                            }



                            sheet.Table.Columns.Add(new WorksheetColumn(50));
                            /*******************************************************************/

                            //** Actual Distribution Columns **//
                            if (SPID == "17")
                            {
                                for (int x = 0; x < (serviceCnt - 5); x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(60));
                                }
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Qty
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Transaction Units
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Unit Price
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Value


                                sheet.Table.Columns.Add(new WorksheetColumn(50));   // Description
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Individuals Served
                                sheet.Table.Columns.Add(new WorksheetColumn(70));   //Demographics Update Required
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Status
                                /***********************************************************************/
                            }
                            else if (SPID == "23")
                            {
                                //11//
                                sheet.Table.Columns.Add(new WorksheetColumn(60));   //Code
                                sheet.Table.Columns.Add(new WorksheetColumn(60));   //Client
                                sheet.Table.Columns.Add(new WorksheetColumn(60));   //Service

                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Qty
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Transaction Units
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Unit Price
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Total Value


                                sheet.Table.Columns.Add(new WorksheetColumn(60));   // Description
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //Status
                                sheet.Table.Columns.Add(new WorksheetColumn(50));   //PG/Seasonal Gift Card Provides
                                sheet.Table.Columns.Add(new WorksheetColumn(60));   //Recipient Full Name 

                                //sheet.Table.Columns.Add(new WorksheetColumn(50));
                                /***********************************************************************/
                            }

                            /** Representative Details**/
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   // Rep Name
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   // Rep Last Name
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   // Rep Phone Number
                            sheet.Table.Columns.Add(new WorksheetColumn(70));   // Rep Secondary Phone Number
                            sheet.Table.Columns.Add(new WorksheetColumn(140));  // Rep Email
                            sheet.Table.Columns.Add(new WorksheetColumn(120));  // Rep Position

                            sheet.Table.Columns.Add(new WorksheetColumn(60));   // Target Date

                            if (SPID != "17")
                            {
                                /** Gender Details**/
                                for (int x = 0; x < genderCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(50));
                                }
                                /** Age Details**/
                                for (int x = 0; x < ageCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(40));
                                }
                                /** Race Details**/
                                for (int x = 0; x < raceCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(110));
                                }
                                /** Ethnicity Details**/
                                for (int x = 0; x < ethinicityCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(100));
                                }
                                /** County Details**/
                                for (int x = 0; x < countyCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(70));
                                }
                                /** City Details**/
                                for (int x = 0; x < cityCnt; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(70));
                                }
                            }


                            if (SPID == "23")
                            {
                                /*Additional Questions*/

                                for (int x = 0; x < 3; x++)
                                {
                                    sheet.Table.Columns.Add(new WorksheetColumn(220));
                                }
                            }

                            /*Questions Detaisl*/

                            for (int x = 0; x < questionsCnt; x++)
                            {
                                sheet.Table.Columns.Add(new WorksheetColumn(220));
                            }

                            #endregion

                            #region TITLE BLOCK 1
                            //string _strTitle = dtgrpReq.Rows[0]["OSERVICES"].ToString();
                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add("Agency Partners Reports - Individuals ", DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle1);
                            cell.MergeAcross = Mergcells;

                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add(_strTitle, DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle2);
                            cell.MergeAcross = Mergcells;

                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlEMPTC);
                            cell.MergeAcross = Mergcells;

                            #endregion

                            #region COLUMNS BLOCK 2
                            excelrow = sheet.Table.Rows.Add();

                            cell = excelrow.Cells.Add("Agency", DataType.String, xAgyHeaderCellstyle);

                            cell = excelrow.Cells.Add("Request No.", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Request Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Completion Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Change Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("Change Operator", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);

                            if (serviceCnt > 0)
                            {
                                cell = excelrow.Cells.Add("REQUEST", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC_sp);
                                cell.MergeAcross = serviceCnt;
                            }

                            if (serviceCnt > 0)
                            {
                                cell = excelrow.Cells.Add("ACTUAL DISTRIBUTION", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC_sp);
                                if (SPID == "17")
                                    cell.MergeAcross = (serviceCnt - 5) + 7;
                                else if (SPID == "23")
                                    cell.MergeAcross = 10;
                            }

                            cell = excelrow.Cells.Add("REPRESENTATIVE DETAILS", DataType.String, oCarlosAg_Excel_Properties.gxlBRCHC_sp);
                            cell.MergeAcross = 5;

                            cell = excelrow.Cells.Add("Target Date", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);

                            if (SPID != "17")
                            {
                                cell = excelrow.Cells.Add("GENDER", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = genderCnt - 1;

                                cell = excelrow.Cells.Add("AGE%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = ageCnt - 1;

                                cell = excelrow.Cells.Add("RACE%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = raceCnt - 1;

                                cell = excelrow.Cells.Add("ETHNICITY%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = ethinicityCnt - 1;

                                cell = excelrow.Cells.Add("COUNTY%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = countyCnt - 1;

                                cell = excelrow.Cells.Add("CITY%", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = cityCnt - 1;
                            }

                            if (SPID == "23")
                            {
                                cell = excelrow.Cells.Add("Additional Questions", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                cell.MergeAcross = 3;
                            }

                            cell = excelrow.Cells.Add(_strTitle + " " + "QUESTIONS", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            if (SPID == "23")
                                cell.MergeAcross = questionsCnt - 5;
                            else
                                cell.MergeAcross = questionsCnt - 1;

                            #endregion

                            #region COLUMNS-2 BLOCK 3
                            excelrow = sheet.Table.Rows.Add();

                            cell = excelrow.Cells.Add("", DataType.String, xAgyHeaderCellstyle);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);

                            //REQUEST
                            if (serviceCnt > 0)
                            {

                                if (SPID == "17")
                                {
                                    foreach (DataRow drserHeads in dtServices.Rows)
                                    {
                                        cell = excelrow.Cells.Add(drserHeads["CAMS_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    }

                                    cell = excelrow.Cells.Add("Participant's First name", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Participant's Last name", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("DOB".ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Client of A Preciosu Child?", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Responsibility for additional documentation", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                }
                                else if (SPID == "23")
                                {
                                    cell = excelrow.Cells.Add("Code", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Children Name", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Children Last Name ", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Service Gifts Idea 1", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Service Gifts Idea 2", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Service Gifts Idea 3", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Clothing/Bike Size (if appl)", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Shoes Size (if app)", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Parent or CG Name ", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Parent or CG Last Name ", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Parent or CG Phone ", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                    cell = excelrow.Cells.Add("Parent or CG Email  ", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                                }
                                cell = excelrow.Cells.Add("Total Ind. Served", DataType.String, oCarlosAg_Excel_Properties.gxlBCHC);
                            }

                            //ACTUAL DISTRIBUTION
                            if (serviceCnt > 0)
                            {
                                if (SPID == "17")
                                {
                                    foreach (DataRow drserHeads in dtServices.Rows)
                                    {
                                        cell = excelrow.Cells.Add(drserHeads["CAMS_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    }
                                }
                                if (SPID == "23")
                                {
                                    cell = excelrow.Cells.Add("Code", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Client", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Service", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);

                                    cell = excelrow.Cells.Add("Total Qty", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Transaction Units", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Unit Price", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Total Value", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Description", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Status", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("PG/Seasonal Gift Card Provides", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Recipient Full Name ", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                }


                                if (SPID == "17")
                                {
                                    cell = excelrow.Cells.Add("Total Qty", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Transaction Units", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Unit Price", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Total Value", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Description", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);

                                    cell = excelrow.Cells.Add("Total Ind. Served", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Demographics Update Require", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    cell = excelrow.Cells.Add("Status", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                }
                                else
                                {
                                    //cell = excelrow.Cells.Add("Status", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                    // cell = excelrow.Cells.Add("Total Ind. Served", DataType.String, oCarlosAg_Excel_Properties.gxlGCHC);
                                }


                            }
                            //REPRESENTATIVE DETAILS
                            cell = excelrow.Cells.Add("Name", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Last Name", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Phone", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Phone2", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Email", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);
                            cell = excelrow.Cells.Add("Position", DataType.String, oCarlosAg_Excel_Properties.gxlBRLHC);

                            //Target Date
                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNLHC);

                            if (SPID != "17")
                            {
                                //GENDER DISTRIBUTION
                                foreach (DataRow drRowsHead in dtGenderData.Rows)
                                {
                                    cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                }

                                //AGE DISTRIBUTION
                                foreach (DataRow drRowsHead in dtAgeData.Rows)
                                {
                                    cell = excelrow.Cells.Add(drRowsHead["id"].ToString() + " yrs.", DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                }

                                //Race DISTRIBUTION
                                foreach (DataRow drRowsHead in dtRaceData.Rows)
                                {
                                    cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                }

                                //Ethnicity DISTRIBUTION
                                foreach (DataRow drRowsHead in dtEthnicityData.Rows)
                                {
                                    cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                }

                                //County DISTRIBUTION
                                foreach (DataRow drRowsHead in dtCountyData.Rows)
                                {
                                    cell = excelrow.Cells.Add(drRowsHead["LookUpDesc"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                }
                                //City DISTRIBUTION
                                foreach (DataRow drRowsHead in dtCityData.Rows)
                                {
                                    cell = excelrow.Cells.Add(drRowsHead["SQR_RESP_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCHC);
                                }

                            }
                            //QUESTIONS 
                            foreach (DataRow drRowsHead in dsQueMast.Tables[0].Rows)
                            {
                                cell = excelrow.Cells.Add(drRowsHead["AGYQ_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNLHC);
                            }

                            #endregion

                            #region ROWS DATA BLOCK4

                            string strAgency = "";
                            string ReqDate = "";
                            string ReqCompDate = "";
                            string strChangeOperator = "";
                            string RequestID = "";
                            string ReqType = "";
                            string ReqSPID = "";
                            string ReqStatus = "";

                            int RowID = 5; int TempRowCount = 0;
                            int chkIsfirst = 0;
                            foreach (DataRow drserDets in dtgrpReq.Rows)
                            {
                                if (chkIsfirst > -1)
                                {
                                    strAgency = drserDets["AGYNAME"].ToString();
                                    ReqDate = Convert.ToDateTime(drserDets["APR_SEREQ_DATE_ADD"]).ToString("MM/dd/yyyy");
                                    ReqCompDate = Convert.ToDateTime(drserDets["APR_SEREQ_DATE_LSTC"]).ToString("MM/dd/yyyy");
                                    strChangeOperator = drserDets["APR_SEREQ_LSTC_OPERATOR"].ToString();
                                    RequestID = drserDets["APR_SEREQ_ID"].ToString();
                                    ReqType = drserDets["APR_SEREQ_SP_TYPE"].ToString();
                                    ReqSPID = drserDets["APR_SEREQ_SP"].ToString();
                                    ReqStatus = drserDets["SERSTATUS"].ToString();

                                    DataTable _dtDemographs = JsonConvert.DeserializeObject<DataSet>(drserDets["APR_SERQ_DEMOGRAPH"].ToString()).Tables[0];
                                    int _INDVMergeDownCount = _dtDemographs.Rows.Count - 1;


                                    excelrow = sheet.Table.Rows.Add();

                                    if (SPID == "17")
                                    {

                                        cell = excelrow.Cells.Add(strAgency, DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                        cell = excelrow.Cells.Add(RequestID, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell = excelrow.Cells.Add(ReqDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell = excelrow.Cells.Add(ReqCompDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell = excelrow.Cells.Add(ReqCompDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell = excelrow.Cells.Add(strChangeOperator, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    }
                                    if (SPID == "23")
                                    {

                                        cell = excelrow.Cells.Add(strAgency, DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                        cell.MergeDown = _INDVMergeDownCount;

                                        cell = excelrow.Cells.Add(RequestID, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _INDVMergeDownCount;

                                        cell = excelrow.Cells.Add(ReqDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _INDVMergeDownCount;

                                        cell = excelrow.Cells.Add(ReqCompDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _INDVMergeDownCount;

                                        cell = excelrow.Cells.Add(ReqCompDate, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _INDVMergeDownCount;

                                        cell = excelrow.Cells.Add(strChangeOperator, DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                    }



                                    ///********************************** --- REQUEST    ---    **************************************///
                                    ///
                                    string str19IndvCount = "";
                                    if (SPID == "17")
                                    {
                                        if (serviceCnt > 0)
                                        {
                                            string[] Servarray = drserDets["APR_SERQ_SERVICES"].ToString().Split(',');
                                            string[] Countarray = drserDets["APR_SERQ_INDIV_CNT"].ToString().Split(',');
                                            if (Servarray.Length > 0)
                                            {
                                                int ttlIndserv = 0;
                                                foreach (DataRow drserHeads in dtServices.Rows)
                                                {
                                                    string strServCode = drserHeads["SP2_CAMS_CODE"].ToString().Trim();
                                                    var serIndex = Array.IndexOf(Servarray, strServCode);
                                                    if (serIndex > -1)
                                                    {
                                                        string serValue = Countarray[serIndex];
                                                        cell = excelrow.Cells.Add(serValue.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                        ttlIndserv += Convert.ToInt32(serValue);
                                                    }
                                                    else
                                                    {
                                                        cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    }
                                                }

                                                if (_dtDemographs.Rows.Count > 0)
                                                {
                                                    cell = excelrow.Cells.Add(_dtDemographs.Rows[0]["FName"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add(_dtDemographs.Rows[0]["LName"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add(_dtDemographs.Rows[0]["DOB"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add(_dtDemographs.Rows[0]["Ispreciouschild"].ToString() == "Y" ? "Yes" : "No", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add("Accepted", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                }
                                                else
                                                {
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                }

                                                cell = excelrow.Cells.Add(ttlIndserv.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC_cr);
                                            }
                                            else
                                            {
                                                foreach (DataRow drserHeads in dtServices.Rows)
                                                {
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                }

                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC);

                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlBCC_cr);
                                            }
                                        }
                                    }
                                    if (SPID == "23")
                                    {
                                        if (_dtDemographs.Rows.Count > 0)
                                        {

                                            int _MiddleRowNo = MiddleRowNumb(_dtDemographs.Rows.Count);
                                            int t = 0;
                                            foreach (DataRow dr23Services in _dtDemographs.Rows)
                                            {
                                                if (t > 0)
                                                {
                                                    excelrow = sheet.Table.Rows.Add();
                                                }

                                                cell = excelrow.Cells.Add(dr23Services["ID"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                if (t > 0)
                                                    cell.Index = 7;
                                                cell = excelrow.Cells.Add(dr23Services["ChildName"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["ChildLName"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["gift1"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["gift2"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["gift3"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["clothsize"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["shoesize"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);

                                                cell = excelrow.Cells.Add(dr23Services["Name"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["Lastname"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["phone"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);
                                                cell = excelrow.Cells.Add(dr23Services["Email"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBCC);


                                                if (t == (_MiddleRowNo))
                                                {
                                                    cell = excelrow.Cells.Add(drserDets["APR_SERQ_INDIV_CNT"].ToString(), DataType.String, gxlBCC_sp_cr);
                                                    //cell.MergeDown = _19SPMergeCount;
                                                    //str23IndvCount = drserDets["APR_SERQ_INDIV_CNT"].ToString();
                                                }
                                                else
                                                    cell = excelrow.Cells.Add("", DataType.String, gxlBCC_sp_cr);
                                                t++;
                                            }


                                        }

                                    }

                                    ///*****************************************************************************************************************///

                                    ///********************************** --- ACTUAL DISTRIBUTION    ---    **************************************///

                                    DataRow[] drReviewDets = dsReviewData.Tables[0].Select("APR_REV_REQ_ID=" + RequestID + "");
                                    if (SPID == "17")
                                    {
                                        if (drReviewDets.Length > 0)
                                        {
                                            DataTable dtReviewDets = drReviewDets.CopyToDataTable();
                                            decimal ttldistributon = 0;
                                            decimal ttlQTY = 0;

                                            foreach (DataRow drserHeads in dtServices.Rows)
                                            {
                                                string strServCode = drserHeads["SP2_CAMS_CODE"].ToString().Trim();
                                                DataRow[] drRevArry = dtReviewDets.Select("APR_REV_SERVICE='" + strServCode + "'");
                                                if (drRevArry.Length > 0)
                                                {
                                                    foreach (DataRow drRDets in drRevArry)
                                                    {
                                                        string Qty = drRDets["APR_REV_QTY"].ToString().Trim();
                                                        cell = excelrow.Cells.Add(Qty.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                        ttlQTY += Qty == "" ? 0 : Convert.ToDecimal(Qty);
                                                    }

                                                }
                                                else
                                                {
                                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                }
                                            }
                                            cell = excelrow.Cells.Add(ttlQTY.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                            string strMeasures = "";
                                            DataTable dtMeasures = JsonConvert.DeserializeObject<DataTable>(Admn_reqreviewandcompletion.WM_GetSPMeasures(dtServices.Rows[0]["SP2_CAMS_CODE"].ToString().Trim(), SPID, "SPMEASURES"));
                                            if (dtMeasures.Rows.Count > 0)
                                            {
                                                strMeasures = dtMeasures.Rows[0]["Measure"].ToString();
                                            }

                                            cell = excelrow.Cells.Add(strMeasures, DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(drReviewDets[0]["APR_REV_UNITPRICE"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(drReviewDets[0]["APR_REV_TOTAL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(drReviewDets[0]["APR_REV_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(drReviewDets[0]["APR_REV_TOT_IND_SERVED"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC_cr);
                                            cell = excelrow.Cells.Add(drReviewDets[0]["APR_REV_DEMOGRAPH_REQ"].ToString() == "Y" ? "Yes" : "No", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(drReviewDets[0]["SERSTATUS"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                        }
                                        else
                                        {
                                            foreach (DataRow drserHeads in dtServices.Rows)
                                            {
                                                cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            }
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC_cr);
                                            cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                            cell = excelrow.Cells.Add(ReqStatus.ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                        }

                                    }


                                    if (SPID == "23")
                                    {

                                        if (drReviewDets.Length > 0)
                                        {
                                            DataTable _dt23RevewData = drReviewDets.CopyToDataTable();
                                            int _MiddleRowNo = MiddleRowNumb(_dtDemographs.Rows.Count);
                                            int i = 0;
                                            for (int x = 0; x < (_dtDemographs.Rows.Count); x++)
                                            {
                                                int _cRowID = (RowID + i);

                                                DataRow[] _dr23RevData = _dt23RevewData.Select("APR_REV_CODE='" + _dtDemographs.Rows[x]["ID"].ToString() + "'");

                                                if (_dr23RevData.Length > 0)
                                                {
                                                    string strServDesc = "";
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_CODE"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_CLIENT"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_ITEM_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_QTY"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_TRANS_UNIT"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_UNITPRICE"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_TOTAL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_DESC"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["SERSTATUS"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_ISGIFTCARD"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add(_dr23RevData[0]["APR_REV_REP"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                                    //if (i == (_MiddleRowNo))
                                                    //{
                                                    //    cell = sheet.Table.Rows[_cRowID].Cells.Add(str19IndvCount, DataType.String, gxlGCC_sp_cr);
                                                    //}
                                                    //else
                                                    //    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, gxlGCC_sp_cr);
                                                }
                                                else
                                                {
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                                    //if (i == (_MiddleRowNo))
                                                    //{
                                                    //    cell = sheet.Table.Rows[_cRowID].Cells.Add(str19IndvCount, DataType.String, gxlGCC_sp_cr);
                                                    //}
                                                    //else
                                                    //    cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, gxlGCC_sp_cr);
                                                }

                                                i++;
                                            }
                                        }
                                        else
                                        {
                                            int i = 0;
                                            for (int x = 0; x < (_dtDemographs.Rows.Count); x++)
                                            {
                                                int _cRowID = (RowID + i);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);
                                                cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlGCC);

                                                i++;
                                            }
                                        }

                                    }

                                    ///*****************************************************************************************************************///

                                    ///****************** SET FIXED EXCEL ROW ************************////
                                    if (SPID == "23")
                                        excelrow = sheet.Table.Rows[RowID];
                                    ////******************************************////

                                    ///********************************************* Representative Details ************************************************///
                                    ///

                                    if (SPID == "17")
                                    {
                                        cell = excelrow.Cells.Add(drserDets["REPFNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["REPLNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["PHONE1"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["PHONE2"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["REPEMAIL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell = excelrow.Cells.Add(drserDets["REPPOST"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                    }
                                    if (SPID == "23")
                                    {
                                        cell = excelrow.Cells.Add(drserDets["REPFNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                        cell = excelrow.Cells.Add(drserDets["REPLNAME"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                        cell = excelrow.Cells.Add(drserDets["PHONE1"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                        cell = excelrow.Cells.Add(drserDets["PHONE2"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                        cell = excelrow.Cells.Add(drserDets["REPEMAIL"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                        cell = excelrow.Cells.Add(drserDets["REPPOST"].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlBRLC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                    }

                                    ///*****************************************************************************************************************///
                                    ///
                                    if (SPID == "17")
                                        cell = excelrow.Cells.Add(Convert.ToDateTime(drserDets["APR_SEREQ_TARGET_DT"].ToString()).ToString("MM/dd/yyyy"), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                    if (SPID == "23")
                                    {
                                        cell = excelrow.Cells.Add(Convert.ToDateTime(drserDets["APR_SEREQ_TARGET_DT"].ToString()).ToString("MM/dd/yyyy"), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                        cell.MergeDown = _INDVMergeDownCount;
                                    }


                                    if (SPID == "23")
                                    {
                                        DataRow[] drDetshead = dt23DemographData.Select("APR_SEREQ_ID='" + RequestID + "'");
                                        if (drDetshead.Length > 0)
                                        {
                                            //cell = excelrow.Cells.Add();
                                            int t = 0;
                                            foreach (DataRow dr in drDetshead)
                                            {
                                                int _cRowID = (RowID + t);
                                                //if (t > 0)
                                                //{
                                                //    excelrow = sheet.Table.Rows.Add();
                                                //}

                                                //if (t > 0)
                                                //    cell = excelrow.Cells.Add();

                                                //cell = sheet.Table.Rows[_cRowID].Cells.Add("Test1", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                //if (t > 0)
                                                //{
                                                //    cell.Index = 38;
                                                //    cell = excelrow.Cells.Add();
                                                //}
                                                //cell = sheet.Table.Rows[_cRowID].Cells.Add(" Test2", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                //if (t > 0)
                                                //    cell = excelrow.Cells.Add();
                                                //cell = sheet.Table.Rows[_cRowID].Cells.Add(" Test3", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                //if (t > 0)
                                                //    cell = excelrow.Cells.Add();
                                                //cell = sheet.Table.Rows[_cRowID].Cells.Add(" Test4", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                //GENDER DISTRIBUTION
                                                int x = 0;
                                                foreach (DataRow drRowsHead in dtGenderData.Rows)
                                                {
                                                    if (t > 0 && x == 1)
                                                        cell.Index = 38;

                                                    if (t > 0)
                                                        cell = excelrow.Cells.Add();

                                                    string code = "gen_" + drRowsHead["Code"].ToString();
                                                    if (dr[code].ToString() != "")
                                                    {
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                        x++;
                                                    }
                                                    else
                                                    {
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                        x++;
                                                    }
                                                }

                                                //excelrow = sheet.Table.Rows.Add();
                                                // AGE DISTRIBUTION
                                                foreach (DataRow drRowsHead in dtAgeData.Rows)
                                                {
                                                    if (t > 0)
                                                        cell = excelrow.Cells.Add();

                                                    string code = "age_" + drRowsHead["id"].ToString();
                                                    // cell = excelrow.Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    if (dr[code].ToString() != "")
                                                    {
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    }
                                                    else
                                                    {
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    }
                                                }

                                                ////Race DISTRIBUTION
                                                foreach (DataRow drRowsHead in dtRaceData.Rows)
                                                {
                                                    if (t > 0)
                                                        cell = excelrow.Cells.Add();

                                                    string code = "rac_" + drRowsHead["code"].ToString();

                                                    if (dr[code].ToString() != "")
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    else
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                }


                                                //Ethnicity DISTRIBUTION
                                                foreach (DataRow drRowsHead in dtEthnicityData.Rows)
                                                {
                                                    if (t > 0)
                                                        cell = excelrow.Cells.Add();

                                                    string code = "eth_" + drRowsHead["code"].ToString();
                                                    // cell = excelrow.Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                    if (dr[code].ToString() != "")
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    else
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                }


                                                //County DISTRIBUTION
                                                foreach (DataRow drRowsHead in dtCountyData.Rows)
                                                {
                                                    if (t > 0)
                                                        cell = excelrow.Cells.Add();

                                                    string code = "cou_" + drRowsHead["code"].ToString();
                                                    //cell = excelrow.Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                    if (dr[code].ToString() != "")
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    else
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                }



                                                //City DISTRIBUTION
                                                foreach (DataRow drRowsHead in dtCityData.Rows)
                                                {
                                                    if (t > 0)
                                                        cell = excelrow.Cells.Add();

                                                    string code = "cit_" + drRowsHead["SQR_RESP_CODE"].ToString().Trim().Replace(" ", "_");
                                                    // cell = excelrow.Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    if (dr[code].ToString() != "")
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add(dr[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNCC);
                                                    else
                                                        cell = sheet.Table.Rows[_cRowID].Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNCC);

                                                }

                                                t++;
                                            }
                                        }
                                    }
                                    // if (SPID == "23")
                                    //  excelrow = sheet.Table.Rows[RowID];

                                    //QUESTIONS DETAILS
                                    foreach (DataRow drRowsQ in dsQueMast.Tables[0].Rows)
                                    {
                                        string code = "que_" + drRowsQ["AGYQ_CODE"].ToString();
                                        if (SPID == "17")
                                            cell = excelrow.Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                        if (SPID == "23")
                                        {
                                            cell = sheet.Table.Rows[RowID].Cells.Add(drserDets[code].ToString(), DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                            cell.MergeDown = _INDVMergeDownCount;
                                        }
                                    }


                                    excelrow = sheet.Table.Rows.Add();
                                    excelrow.Height = 5;
                                    cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlNLC);
                                    cell.MergeAcross = Mergcells;

                                    ///**************************  SET ROWID *********************************////

                                    if (chkIsfirst == 0)
                                    {
                                        TempRowCount = RowID + _INDVMergeDownCount + 1 + 1;
                                    }
                                    else if (chkIsfirst > 0)
                                    {

                                        TempRowCount = (TempRowCount + _INDVMergeDownCount + 1) + 1;
                                    }



                                    //**************************************************************************///
                                }
                                RowID = (TempRowCount == 0 ? 5 : TempRowCount);
                                chkIsfirst++;
                            }

                            #endregion

                        }
                        #endregion
                    }
                    else
                    {
                        if (SPgrpType == "G")
                        {
                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add("Agency Partners Reports - Groups ", DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle1);
                            cell.MergeAcross = Mergcells;
                        }
                        else
                        {
                            excelrow = sheet.Table.Rows.Add();
                            cell = excelrow.Cells.Add("Agency Partners Reports - Individuals ", DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle1);
                            cell.MergeAcross = Mergcells;
                        }

                        excelrow = sheet.Table.Rows.Add();
                        cell = excelrow.Cells.Add(_strTitle, DataType.String, oCarlosAg_Excel_Properties.gxlTitle_CellStyle2);
                        cell.MergeAcross = Mergcells;

                        excelrow = sheet.Table.Rows.Add();
                        cell = excelrow.Cells.Add("", DataType.String, oCarlosAg_Excel_Properties.gxlEMPTC);
                        cell.MergeAcross = Mergcells;

                        excelrow = sheet.Table.Rows.Add();
                        cell = excelrow.Cells.Add("No Records Found!", DataType.String, oCarlosAg_Excel_Properties.gxlERRMSG);
                        cell.MergeAcross = Mergcells;

                    }


                }
                /**************************/
                #region excelfile Creation
                FileStream stream = new FileStream(Server.MapPath("~/EXCEL/Service_Request_Report.xls"), FileMode.Create);

                book.Save(stream);
                stream.Close();

                FileInfo file = new FileInfo(Server.MapPath("~/EXCEL/Service_Request_Report.xls"));
                if (file.Exists)
                {
                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=Service_Request_Report_" + DateTime.Now.ToShortDateString() + ".xls");
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", file.Length.ToString());
                    Response.WriteFile(file.FullName);
                    Response.End();
                }
                #endregion
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
            }
        }

        public List<params1> WM_DecryptUrl(string ourl)
        {
            List<params1> result = new List<params1>();
            try
            {
                //ourl = ourl.Remove(0, 1);
                // ourl = ourl.Remove(ourl.Length - 1);

                string _decrypturl = AgencyPartnerPortalDB.DatabaseLayer.Capsystems.Decrypt(HttpUtility.UrlDecode(ourl));

                JavaScriptSerializer json = new JavaScriptSerializer();
                List<params1> objArray = json.Deserialize<List<params1>>(_decrypturl);

                result = objArray;

            }
            catch (Exception ex)
            {
                //result = ex.Message;
            }
            return result;
        }

        public DataTable BuildDataSetWithAllTables(DataTable dtTemp, DataTable dtQuesAns, DataSet _dsMasters, string _SPID, string _strGrpType)
        {

            DataTable dtResult = new DataTable();
            dtResult = dtTemp.Copy();



            DataTable dtQueData = new DataTable();
            DataSet dsQueMast = AgencyPartnerPortalDB.DatabaseLayer.Capsystems.CAPS_AGCYSERQASSOC_GET(_SPID, "", "SERQASSOC");
            if (dsQueMast.Tables.Count > 0)
            {
                dtQueData = dsQueMast.Tables[0].Copy();
            }
            dtResult = AddPivotColumns(dtResult, dtQueData, "AGYQ_CODE", "que");

            if (_strGrpType == "G")
            {

                /* Masters Table */
                DataTable dtGenderData = _dsMasters.Tables["Gender"];
                DataTable dtAgeData = CommonFiles.Commonfunctions.BuildAgeMaster(_SPID);
                DataTable dtRaceData = _dsMasters.Tables["Race"];
                DataTable dtEthnicityData = _dsMasters.Tables["Ethnicity"];
                DataTable dtCountyData = _dsMasters.Tables["County"];
                DataTable dtCityData = _dsMasters.Tables["City"];
                DataTable dtLanguagesData = _dsMasters.Tables["Languages"];

                /*Add Pivot Columns*/

                dtResult = AddPivotColumns(dtResult, dtGenderData, "Code", "gen");
                dtResult = AddPivotColumns(dtResult, dtAgeData, "id", "age");
                dtResult = AddPivotColumns(dtResult, dtRaceData, "Code", "rac");
                dtResult = AddPivotColumns(dtResult, dtEthnicityData, "Code", "eth");
                dtResult = AddPivotColumns(dtResult, dtCountyData, "Code", "cou");
                dtResult = AddPivotColumns(dtResult, dtCityData, "SQR_RESP_CODE", "cit");
                dtResult = AddPivotColumns(dtResult, dtLanguagesData, "Code", "lan");
            }

            foreach (DataRow dr in dtResult.Rows)
            {
                string strReqID = dr["APR_SEREQ_ID"].ToString();

                DataRow[] drQA = dtQuesAns.Select("APRQUES_REQ_ID=" + strReqID + "");

                DataSet _dtDemographTbls = JsonConvert.DeserializeObject<DataSet>(dr["APR_SERQ_DEMOGRAPH"].ToString());
                if (_dtDemographTbls.Tables.Count > 0)
                {
                    if (_strGrpType == "G")
                    {
                        //  DataTable _dtServiceDetsTable = ConvertToPivotTable(_dtDemographTbls.Tables[0].Copy(), "srv");
                        DataTable _dtGenderTable = _dtDemographTbls.Tables[1].Copy();
                        DataTable _dtAgeTable = _dtDemographTbls.Tables[2].Copy();
                        DataTable _dtRaceTable = _dtDemographTbls.Tables[3].Copy();
                        DataTable _dtEthencityTable = _dtDemographTbls.Tables[4].Copy();
                        DataTable _dtCountyTable = _dtDemographTbls.Tables[5].Copy();
                        DataTable _dtCityTable = _dtDemographTbls.Tables[6].Copy();

                        /*Gender Table Add Rows to Pivot Table*/
                        foreach (DataRow drRow in _dtGenderTable.Rows)
                        {
                            dr["gen_" + drRow["code"]] = drRow["Value"].ToString();
                        }

                        /*Age Table Add Rows to Pivot Table*/
                        foreach (DataRow drRow in _dtAgeTable.Rows)
                        {
                            dr["age_" + drRow["Code"]] = drRow["Value"].ToString();
                        }

                        /*Race Table Add Rows to Pivot Table*/
                        foreach (DataRow drRow in _dtRaceTable.Rows)
                        {
                            dr["rac_" + drRow["Code"]] = drRow["Value"].ToString();
                        }

                        /*Ethenicity Table Add Rows to Pivot Table*/
                        foreach (DataRow drRow in _dtEthencityTable.Rows)
                        {
                            dr["eth_" + drRow["Code"]] = drRow["Value"].ToString();
                        }

                        /*County Table Add Rows to Pivot Table*/
                        foreach (DataRow drRow in _dtCountyTable.Rows)
                        {
                            dr["cou_" + drRow["Code"]] = drRow["Value"].ToString();
                        }

                        /*City Table Add Rows to Pivot Table*/
                        foreach (DataRow drRow in _dtCityTable.Rows)
                        {
                            dr["cit_" + drRow["Code"].ToString().Trim().Replace(" ", "_")] = drRow["Value"].ToString();
                        }
                    }

                    /*Question with Answers */
                    DataTable odtQuesAns = new DataTable();
                    if (drQA.Length > 0)
                    {
                        odtQuesAns = drQA.CopyToDataTable();

                        foreach (DataRow drRow in odtQuesAns.Rows)
                        {
                            dr["que_" + drRow["APR_SEREQ_QCODE"].ToString()] = drRow["APR_SEREQ_RESPONSE"].ToString();
                        }
                    }


                }

            }

            return dtResult;

        }

        public DataTable BuildDataSet23WithAllTables(DataTable dtTemp, DataTable dtQuesAns, DataSet _dsMasters, string _SPID, string _strGrpType)
        {
            DataTable dtResult = new DataTable();
            //dtResult = dtTemp.Copy();

            DataTable dt23Data = new DataTable();

            int i = 0;
            foreach (DataRow dr in dtTemp.Rows)
            {
                string strReqID = dr["APR_SEREQ_ID"].ToString();

                DataRow[] drQA = dtQuesAns.Select("APRQUES_REQ_ID=" + strReqID + "");

                DataSet _dtDemographTbls = JsonConvert.DeserializeObject<DataSet>(dr["APR_SERQ_DEMOGRAPH"].ToString());
                _dtDemographTbls.Tables[0].Columns.Add("APR_SEREQ_ID", typeof(System.String));
                if (_dtDemographTbls.Tables.Count > 0)
                {
                    if (i == 0)
                    {
                        if ((_strGrpType == "I" && _SPID == "23"))
                        {

                            dt23Data = _dtDemographTbls.Tables[0].Copy();
                            /* Masters Table */
                            DataTable dtGenderData = _dsMasters.Tables["Gender"];
                            DataTable dtAgeData = CommonFiles.Commonfunctions.BuildAgeMaster(_SPID);
                            DataTable dtRaceData = _dsMasters.Tables["Race"];
                            DataTable dtEthnicityData = _dsMasters.Tables["Ethnicity"];
                            DataTable dtCountyData = _dsMasters.Tables["County"];
                            DataTable dtCityData = _dsMasters.Tables["City"];
                            DataTable dtLanguagesData = _dsMasters.Tables["Languages"];

                            /*Add Pivot Columns*/

                            dt23Data = AddPivotColumns(dt23Data, dtGenderData, "Code", "gen");
                            dt23Data = AddPivotColumns(dt23Data, dtAgeData, "id", "age");
                            dt23Data = AddPivotColumns(dt23Data, dtRaceData, "Code", "rac");
                            dt23Data = AddPivotColumns(dt23Data, dtEthnicityData, "Code", "eth");
                            dt23Data = AddPivotColumns(dt23Data, dtCountyData, "Code", "cou");
                            dt23Data = AddPivotColumns(dt23Data, dtCityData, "SQR_RESP_CODE", "cit");
                            dt23Data = AddPivotColumns(dt23Data, dtLanguagesData, "Code", "lan");


                            foreach (DataRow _dr in dt23Data.Rows)
                            {
                                _dr["APR_SEREQ_ID"] = strReqID;
                            }
                        }
                    }
                    if (i > 0)
                    {
                        foreach (DataRow _dr in _dtDemographTbls.Tables[0].Rows)
                        {
                            _dr["APR_SEREQ_ID"] = strReqID;
                            dt23Data.ImportRow(_dr);

                        }
                    }
                }

                i++;
            }

            foreach (DataRow _dr in dt23Data.Rows)
            {

                string strGValue = _dr["Gender"].ToString();
                _dr["gen_" + strGValue] = "1";


                string strAgValue = _dr["Age"].ToString().Remove((_dr["Age"].ToString().Length - 5));
                _dr["age_" + strAgValue] = "1";

                string strRcValue = _dr["Racecd"].ToString();
                _dr["rac_" + strRcValue] = "1";

                string strEthValue = _dr["Ethenicitycd"].ToString();
                _dr["eth_" + strEthValue] = "1";

                string strCityValue = _dr["City"].ToString();
                _dr["cit_" + strCityValue.ToString().Trim().Replace(" ", "_")] = "1";

                string strCounValue = _dr["County"].ToString();
                _dr["cou_" + strCounValue] = "1";
            }



            dtResult = dt23Data;

            return dtResult;
        }

        DataTable AddPivotColumns(DataTable _Maintbl, DataTable _pvtTable, string _ColumnName, string _colTyp)
        {
            DataTable dtRes = new DataTable();
            try
            {
                if (_colTyp == "que")
                {
                    int[] uniqueAncType = _pvtTable.AsEnumerable().Select(x => x.Field<int>(_ColumnName.ToString())).Distinct().ToArray();
                    foreach (int col in uniqueAncType)
                    {
                        _Maintbl.Columns.Add(_colTyp + "_" + col.ToString(), typeof(string));
                    }
                }
                else
                {
                    string[] uniqueAncType = _pvtTable.AsEnumerable().Select(x => x.Field<string>(_ColumnName.ToString())).Distinct().ToArray();
                    foreach (string col in uniqueAncType)
                    {
                        _Maintbl.Columns.Add(_colTyp + "_" + col.Trim().Replace(" ", "_"), typeof(string));
                    }
                }
                dtRes = _Maintbl.Copy();
            }
            catch (Exception ex)
            {
            }
            return dtRes;
        }

        DataTable ConvertToPivotTable(DataTable _dtTemp, string _colTyp)
        {
            DataTable dtRes = new DataTable();
            try
            {

                if (_dtTemp.Rows.Count > 0)
                {
                    string[] uniqueAncType = _dtTemp.AsEnumerable().Select(x => x.Field<string>("Code")).Distinct().ToArray();

                    DataTable pivot = new DataTable();
                    foreach (string col in uniqueAncType)
                    {
                        pivot.Columns.Add(_colTyp + "_" + col, typeof(string));
                    }

                    DataRow newRow = pivot.Rows.Add();
                    foreach (DataRow ancType in _dtTemp.Rows)
                    {
                        newRow[_colTyp + "_" + ancType.Field<string>("Code")] = ancType.Field<string>("Value");
                    }

                    // var groups = _dtTemp.AsEnumerable()
                    //.GroupBy(x => new {
                    //    opr_dt = x.Field<DateTime>("OPR_DT"),
                    //    opr_hr = x.Field<int>("OPR_HR"),
                    //    anc_region = x.Field<string>("ANC_REGION"),
                    //    run_id = x.Field<string>("MARKET_RUN_ID")
                    //}).ToList();

                    //var groups = _dtTemp.AsEnumerable()
                    //  .GroupBy(x => new {
                    //      value = x.Field<string>("Value")
                    //  }).ToList();

                    //foreach (var group in groups)
                    //{
                    //    DataRow newRow = pivot.Rows.Add();
                    //    newRow["OPR_DT"] = group.Key.value;
                    //    newRow["OPR_HR"] = group.Key.opr_hr;
                    //    newRow["ANC_REGION"] = group.Key.anc_region;
                    //    newRow["MARKET_RUN_ID"] = group.Key.run_id;
                    //    foreach (DataRow ancType in group)
                    //    {
                    //        newRow[ancType.Field<string>("ANC_TYPE")] = ancType.Field<decimal>("MW");
                    //    }


                    //}
                    dtRes = pivot.Copy();
                }

            }
            catch (Exception ex)
            {

            }
            return dtRes;
        }

        public int MiddleRowNumb(int RowCount)
        {
            int resNumb = 0;
            try
            {
                decimal x = (RowCount / 2);
                int Roundup = Convert.ToInt32(Math.Round(x, MidpointRounding.AwayFromZero));
                resNumb = Roundup;
            }
            catch
            {

            }

            return resNumb;
        }

    }

}

