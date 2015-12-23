using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SysFreight.Components;
using System.IO;
using static SysFreight.Components.Modfunction;
using System.Collections;
using System.Data.OleDb;
using System.Management;
using System.Net;

namespace EGDC_Trendz_EDI
{
    public partial class frmsTrendz : SysFreight.Components.frmMainBase
    {
        #region Delcare
        string strLocalConnect = "";
        string m_strExportDatabase = "";
        string m_strExportWebSite = "";
        string m_strExportUserID = "";
        string m_strExportPassword = "";
        string strSite = "";
        string UserID = "";
        string Password = "";
        string WebSite = "";
        string m_EDIPath = "";
        string m_LogPath = "";
        string m_BackupEDIPath = "";
        string m_BackupLogPath = "";
        string[] args = null;
        string DateFormat = "";
        string Language = "";
        string m_strYear = "";
        string m_strMth = "";
        string m_strDay = "";
        string m_strHour = "";
        string m_strMin = "";
        string m_strSecond = "";
        DataLayer m_DataLayer = new DataLayer();
        #endregion

        #region EDI AND SysFreight CommonFunction ALL
        #region SysfreightHelpFunction
        private string CheckNull(object ojb)
        {
            short intDateType = 0;
            return Modfunction.CheckNull(ojb, ref intDateType).ToString();
        }

        private object CheckNull(object ojb, short intDateType)
        {
            return Modfunction.CheckNull(ojb, ref intDateType);
        }

        private DateTime CheckNullDate(object ojb, int intDateType)
        {
            short intType = 2;
            return Convert.ToDateTime(Modfunction.CheckNull(ojb, ref intType));
        }

        private Double CheckNullDouble(object ojb, int intDateType)
        {
            short intType = 2;
            return Convert.ToDouble(Modfunction.CheckNull(ojb, ref intType));
        }

        private Int32 CheckNullInt(object ojb, int intDateType)
        {
            short intType = 2;
            return Convert.ToInt32(Modfunction.CheckNull(ojb, ref intType));
        }

        private string ReplaceWithNull(object ojb)
        {
            short intDateType = 0;
            return Modfunction.ReplaceWithNull(ref ojb, ref intDateType);
        }

        private string ReplaceWithNull(object ojb, short intDateType)
        {
            return Modfunction.ReplaceWithNull(ref ojb, ref intDateType);
        }

        private string ReplaceWithNull(object ojb, int intDateType)
        {
            short intType = 0;
            if (intDateType == 1) { intType = 1; }
            if (intDateType == 2) { intType =2; }
            return Modfunction.ReplaceWithNull(ref ojb, ref intType);
        }
        #endregion

        #region SaedFunction
        string m_strFolder = "";
        string m_strOUTFolder = "";
        string int_SaedTrxNo = "";
        string strSchedulerFlag = "";
        string m_strFilter1 = "";
        string strEdiName = "";
        int intLineItemNo = -1;
        int intStarLineItemNo = 0;
        private void getFolderName()
        {
            DataTable dtRec = GetSQLCommandReturnDTNew("Select TrxNo,InFolder,OutFolder,SchedulerFlag,Filter1,EdiName from Saed1 Where EdiName = 'EGDC DELIVERY ORDER'");
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                m_strFolder = CheckNull(dtRec.Rows[0]["InFolder"]);
                m_strOUTFolder = CheckNull(dtRec.Rows[0]["OutFolder"]);
                int_SaedTrxNo = CheckNull(dtRec.Rows[0]["TrxNo"], 1).ToString();
                strSchedulerFlag = CheckNull(dtRec.Rows[0]["SchedulerFlag"]);
                m_strFilter1 = CheckNull(dtRec.Rows[0]["Filter1"]);
                strEdiName = CheckNull(dtRec.Rows[0]["EdiName"]);
            }
        }

        private void saveSaed2(string FileName, string JobNo, string RefNo1, string RefType1, string Remark, string MsgType)
        {
            if (intLineItemNo == -1)
            {
                DataTable dtRec = GetSQLCommandReturnDTNew("Select Max(LineItemNo) from saed2 Where TrxNo = " + this.int_SaedTrxNo.ToString());
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    intLineItemNo = CheckNullInt(dtRec.Rows[0][0], 1)+ 1;
                }
                else
                {
                    intLineItemNo = 1;
                }
            }
            if (intStarLineItemNo == -1) { intStarLineItemNo = intLineItemNo; }
            GetSQLCommandReturnIntNew("Insert into Saed2(TrxNo,LineItemNo,Type,CreateBy,CreateDateTime,FileName,JobNo,RefNo1,RefType1,Remark,MsgType,EdiName) Values(" + int_SaedTrxNo.ToString() + "," + intLineItemNo.ToString() + ",'1'," + ReplaceWithNull(g_strUserID) + ",getdate()," + ReplaceWithNull(FileName) + "," + ReplaceWithNull(JobNo) + "," + ReplaceWithNull(RefNo1) + "," + ReplaceWithNull(RefType1) + "," + ReplaceWithNull(Remark) + "," + ReplaceWithNull(MsgType) + "," + ReplaceWithNull(strEdiName) + ")");
            intLineItemNo = intLineItemNo + 1;
        }

        private string[] getFileList(string strFileType)
        {
            List<string> list = new List<string>();
            list.Clear();
            if (m_strFolder.Trim() == "") { return null; }
            if (!Directory.Exists(@m_strFolder)) { return null; }
            DirectoryInfo Dir = new DirectoryInfo(@m_strFolder);
            foreach (FileInfo FI in Dir.GetFiles())
            {
                if (System.IO.Path.GetExtension(FI.Name) == "." + strFileType)
                {
                    list.Add(FI.Name);
                }
            }
            return list.ToArray();
        }
        #endregion

        #region CheckAndLoginFunction
        private Boolean checkDatabaseLogin()
        {
            DataTable dtRec = GetSQLCommandReturnDTNew("Select 1");
            if (dtRec == null || dtRec.Rows.Count == 0)
            {
                if (m_LogPath.Trim() == "") { m_LogPath = @Directory.GetCurrentDirectory().Trim() + @"\Log"; }
                if (!Directory.Exists(@m_LogPath)) { Directory.CreateDirectory(@m_LogPath); }
                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                if (!Directory.Exists(@m_BackupLogPath)) { Directory.CreateDirectory(@m_BackupLogPath); }
                string strWriteLog = "";
                string strRowDelimiter = "\r\n";
                strWriteLog = "";
                strWriteLog = strWriteLog + "Run Date And Time are " + m_strDay + @"\" + m_strMth + @"\" + m_strYear + " " + m_strHour + @":" + m_strMin + @":" + m_strSecond + strRowDelimiter;
                strWriteLog = strWriteLog + "Can not connect the database" + strRowDelimiter;
                string strFileName = "RunProject_" + m_strYear.Substring(2, 2) + m_strMth + m_strDay + m_strHour + m_strMin + m_strSecond;
                StreamWriter wBackupLog;
                StreamWriter wLog;
                wLog = File.CreateText(@m_LogPath + @"\" + @strFileName + @".LOG");
                wBackupLog = File.CreateText(@m_BackupLogPath + @"\" + @strFileName + @".LOG");
                wLog.Write(strWriteLog);
                wBackupLog.Write(strWriteLog);
                wLog.Close();
                wBackupLog.Close();
                return false;
            }
            return true;
        }

        private void getLoginInfo()
        {
            string sIniFile = "";
            sIniFile = System.IO.Directory.GetCurrentDirectory().Trim() + @"\Setting.ini";
            string SysfreightSetting = "SysfreightSetting";
            string strThree = "Server";
            string strFore = "Localhost";
            m_strExportWebSite = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strThree = "Database";
            strFore = "dmoFreight";
            m_strExportDatabase = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strThree = "UserId";
            strFore = "Sa";
            m_strExportUserID = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strThree = "Password";
            strFore = "";
            m_strExportPassword = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strLocalConnect = "data source=" + m_strExportWebSite + ";initial catalog=" + m_strExportDatabase + ";user id=" + m_strExportUserID + ";password=" + m_strExportPassword + ";persist security info=False";
            //SysfreightSetting = "AutoCountSetting";
            //strThree = "Server";
            //strFore = "Localhost";
            //m_strAutoCountWebSite = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            //strThree = "Database";
            //strFore = "dmoFreight";
            //m_stAutoCounttDatabase = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            //strThree = "UserId";
            //strFore = "Sa";
            //m_strAutoCountDatabaseUserID = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            //strThree = "Password";
            //strFore = "";
            //m_stAutoCountDatabasePassword = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            //strThree = "AutoCountUserId";
            //strFore = "admin";
            //m_strAutoCountUserID = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            //strThree = "AutoCountPassword";
            //strFore = "admin";
            //m_stAutoCountPassword = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            Modfunction.strSqlConn = strLocalConnect;
            Modfunction.strSqlConn2 = strLocalConnect;
            SysfreightSetting = "PathSetting";
            strThree = "EDIPath";
            strFore = "";
            m_EDIPath = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strThree = "LogPath";
            strFore = "";
            m_LogPath = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strThree = "EDIBackupPath";
            strFore = "";
            m_BackupEDIPath = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            strThree = "LogBackupPath";
            strFore = "";
            m_BackupLogPath = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
        }

        #endregion

        #region SQlUpdateFunction
        private Boolean connect(string strModle)
        {
            DataTable dtRec = null;
            if (strModle == "E")
            {
                g_strCurrentDatabase = this.m_strExportDatabase;
                string strhttp = m_strExportWebSite.Trim().Substring(0, "http:".Length).ToLower();
                string strSysWS = m_strExportWebSite.Trim().Substring(m_strExportWebSite.Trim().Length - "SysFreightWS/".Length, "SysFreightWS/".Length).ToLower();
                if (strhttp == "http:" && strSysWS == "SysFreightWS/".ToLower())
                {
                    g_WebUrl = this.m_strExportWebSite.Trim();
                }
                else
                {
                    g_WebUrl = "Http://" + this.m_strExportWebSite.Trim() + "/SysFreightWS/";
                }
                CurrentServiceAndDatabase = g_WebUrl + "," + this.m_strExportDatabase;
                g_strUserID = this.m_strExportUserID;
                g_strPassword = this.m_strExportPassword;
                strSite = this.m_strExportDatabase;
            }
            UserID = g_strUserID;
            Password = g_strPassword;
            if (UserID.Trim() != string.Empty)
            {
                CurrentServiceAndDatabase = g_WebUrl + "," + g_strCurrentDatabase;
                Boolean blnError = false;
                try
                {
                    GetConnectionStringValue(strSite);
                }
                catch (Exception ex)
                {
                    blnError = true;
                    ex.Data.Clear();
                }
                if (blnError)
                {
                    MessageBox.Show("Unable to connect to the remote server", "Error Web Server", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
                g_strCurrentDatabase = strSite;
                CurrentServiceAndDatabase = g_WebUrl + "," + g_strCurrentDatabase;
                dtRec = GetUserInfoNoPassword(UserID);
                if (dtRec == null || dtRec.Rows.Count == 0)
                {
                    MessageBoxShow(this, "msg3", "Please make sure you enter a correct userid!", "Wrong UserID");
                    return false;
                }
                DataLayerResult dlResult;
                dlResult = m_DataLayer.Login(UserID, Password);
                if (dlResult == DataLayerResult.Success)
                {
                    return true;
                }
                else if (dlResult == DataLayerResult.ConnectionFailure)
                {
                    return false;
                }
                else
                {
                    MessageBoxShow(this, "msg1", "Please make sure you enter a correct password!", "Wrong password");
                }
            }
            else if (UserID.Trim() == String.Empty)
            {
                MessageBoxShow(this, "msg2", "User name and Password must Input", "Information!");
            }
            else if (WebSite == "")
            {
                MessageBoxShow(this, "msg2", "web site must Input", "Information!");
            }
            else if (strSite == "")
            { MessageBoxShow(this, "msg2", "database must Input", "Information!"); }
            return false;
        }

        private DataTable GetSQLCommandReturnDTNew(string SqlCommand)
        {
            if (this.strLocalConnect == "")
            {
                return GetSQLCommandReturnDT(SqlCommand);
            }
            else
            {
                return GetSQLCommandReturnDTSysAdmin(SqlCommand);
            }
        }

        private int GetSQLCommandReturnIntNew(string SqlCommand)
        {
            if (this.strLocalConnect == "")
            {
                return GetSQLCommandReturnInt(SqlCommand);
            }
            else
            {
                int intEx = 0;
                return GetSQLCommandReturnIntSysAdmin(SqlCommand, ref intEx);
            }
        }

        #endregion

        private void Main()
        {
            if (args.Length >= 2)
                if (args.Length == 2)
                {
                    string[] ConnectionInfo = args[1].Split(',');
                    if (ConnectionInfo.Length >= 4)
                    {
                        WebSite = ConnectionInfo[0];
                        if (WebSite.Substring(0, 1) == "(") { WebSite = WebSite.Substring(1, WebSite.Length - 1); }
                        if (WebSite.Substring(WebSite.Length - 1, 1) == "(") { WebSite = WebSite.Substring(0, WebSite.Length - 1); }

                        strSite = ConnectionInfo[1];
                        DateFormat = ConnectionInfo[2];
                        Language = ConnectionInfo[3];
                        UserID = ConnectionInfo[4];
                        Password = ConnectionInfo[5];
                        if (Language == "ENU")
                        { g_strActiveLanguageType = "en-US"; }
                        else if (Language == "CHS")
                        { g_strActiveLanguageType = "zh-CN"; }
                        if (g_strDateFormat == "dd/MM/yyyy")
                        {
                            CSHORT_DATE_PATTERN = "dd-MM-yy";
                            CLONG_DATE_PATTERN = "dd-MMM-yy";
                            CSHORT_LDATE_PATTERN = "dd-MM-yyyy";
                            CSHORT_LDATE_PATTERNMask = "00-00-0000";
                            CLONG_LDATE_PATTERN = "dd-MMM-yyyy";
                            CSHORT_DATETIME_PATTERN = "dd-MM-yy HH:mm";
                            CLONG_DATETIME_PATTERN = "dd-MMM-yy HH:mm";
                            CSHORT_LDATETIME_PATTERN = "dd-MM-yyyy HH:mm";
                            CLONG_LDATETIME_PATTERN = "dd-MMM-yyyy HH:mm";
                            CLONG_LDATETIMESS_PATTERN = "dd-MMM-yyyy HH:mm:ss";
                            CNULL_DATE_PATTERN = "dd/MM/yyyy";
                            CNULL_SHORTDATE_PATTERN = "dd/MM/yy";
                            CNULL_DATE = "31/12/1899";
                            CNULL_DATE1 = "31-12-1899";
                        }
                        else if (g_strDateFormat == "yyyy/MM/dd")
                        {
                            CSHORT_DATE_PATTERN = "yy-MM-dd";
                            CLONG_DATE_PATTERN = "yy-MMM-dd";
                            CSHORT_LDATE_PATTERN = "yyyy-MM-dd";
                            CSHORT_LDATE_PATTERNMask = "0000-00-00";
                            CLONG_LDATE_PATTERN = "yyyy-MMM-dd";
                            CSHORT_DATETIME_PATTERN = "yy-MM-dd HH:mm";
                            CLONG_DATETIME_PATTERN = "yy-MMM-dd HH:mm";
                            CSHORT_LDATETIME_PATTERN = "yyyy-MM-dd HH:mm";
                            CLONG_LDATETIME_PATTERN = "yyyy-MMM-dd HH:mm";
                            CLONG_LDATETIMESS_PATTERN = "yyyy-MM-dd HH:mm:ss";
                            CNULL_DATE_PATTERN = "yyyy/MM/dd";
                            CNULL_SHORTDATE_PATTERN = "yy/MM/dd";
                            CNULL_DATE = "1899/12/31";
                            CNULL_DATE1 = "1899-12-31";
                        }
                        else if (g_strDateFormat == "MM/dd/yyyy")
                        {

                            CSHORT_DATE_PATTERN = "MM-dd-yy";
                            CLONG_DATE_PATTERN = "MMM-dd-yy";
                            CSHORT_LDATE_PATTERN = "MM-dd-yyyy";
                            CSHORT_LDATE_PATTERNMask = "00-00-0000";
                            CLONG_LDATE_PATTERN = "MMM-dd-yyyy";
                            CSHORT_DATETIME_PATTERN = "MM-dd-yy HH:mm";
                            CLONG_DATETIME_PATTERN = "MMM-dd-yy HH:mm";
                            CSHORT_LDATETIME_PATTERN = "MM-dd-yyyy HH:mm";
                            CLONG_LDATETIME_PATTERN = "MMM-dd-yyyy HH:mm";
                            CLONG_LDATETIMESS_PATTERN = "MMM-dd-yyyy HH:mm:ss";
                            CNULL_DATE_PATTERN = "MM/dd/yyyy";
                            CNULL_SHORTDATE_PATTERN = "MM/dd/yy";
                            CNULL_DATE = "12/31/1899";
                            CNULL_DATE1 = "12-31-1899";
                        }
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(g_strActiveLanguageType);
                        System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = DateFormat;
                    }
                }

                else
                {
                    string strCon = "";
                    strCon = args[1];
                    for (int intI = 1; intI < args.Length; intI++)
                    { strCon = strCon + " " + args[intI]; }
                    if (strCon.Length > 0)
                    {
                        strLocalConnect = strCon.Substring(0, strCon.LastIndexOf(";"));
                        string[] str = strCon.Split(';');
                        if (str != null && str.Length == 15)
                        {
                            strLocalConnect = "data source=" + str[4].Split('=')[1] + ";initial catalog=" + str[7].Split('=')[1] + ";user id=" + str[5].Split('=')[1] + ";password=" + str[6].Split('=')[1] + ";persist security info=False";
                        }
                        Modfunction.strSqlConn = strLocalConnect;
                        Modfunction.strSqlConn2 = strLocalConnect;
                    }
                }
            else
            { this.getLoginInfo(); }
            string[] strFormat = System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern.Split('/');
            int intY = 0, intD = 0, intM = 0;
            for (int intI = 0; intI < strFormat.Length; intI++)
            {
                if (strFormat[intI].ToLower().IndexOf("y") > 0)
                { intY = intI; }
                else if (strFormat[intI].ToLower().IndexOf("d") > 0)
                { intD = intI; }
                else if (strFormat[intI].ToLower().IndexOf("m") > 0)
                { intM = intI; }
            }
            if (intY == 0)
            {
                if (intD != 1) { g_strDateFormat = "yyyy/MM/dd"; }
            }

            else if (intY == 2)
            {
                if (intD == 0)
                { g_strDateFormat = "dd/MM/yyyy"; }
                else
                { g_strDateFormat = "MM/dd/yyyy"; }
            }
            if (g_strDateFormat == "dd/MM/yyyy")
            {
                CSHORT_DATE_PATTERN = "dd-MM-yy";
                CLONG_DATE_PATTERN = "dd-MMM-yy";
                CSHORT_LDATE_PATTERN = "dd-MM-yyyy";
                CSHORT_LDATE_PATTERNMask = "00-00-0000";
                CLONG_LDATE_PATTERN = "dd-MMM-yyyy";
                CSHORT_DATETIME_PATTERN = "dd-MM-yy HH:mm";
                CLONG_DATETIME_PATTERN = "dd-MMM-yy HH:mm";
                CSHORT_LDATETIME_PATTERN = "dd-MM-yyyy HH:mm";
                CLONG_LDATETIME_PATTERN = "dd-MMM-yyyy HH:mm";
                CLONG_LDATETIMESS_PATTERN = "dd-MMM-yyyy HH:mm:ss";
                CNULL_DATE_PATTERN = "dd/MM/yyyy";
                CNULL_SHORTDATE_PATTERN = "dd/MM/yy";
                CNULL_DATE = "31/12/1899";
                CNULL_DATE1 = "31-12-1899";
            }
            else if (g_strDateFormat == "yyyy/MM/dd")
            {
                CSHORT_DATE_PATTERN = "yy-MM-dd";
                CLONG_DATE_PATTERN = "yy-MMM-dd";
                CSHORT_LDATE_PATTERN = "yyyy-MM-dd";
                CSHORT_LDATE_PATTERNMask = "0000-00-00";
                CLONG_LDATE_PATTERN = "yyyy-MMM-dd";
                CSHORT_DATETIME_PATTERN = "yy-MM-dd HH:mm";
                CLONG_DATETIME_PATTERN = "yy-MMM-dd HH:mm";
                CSHORT_LDATETIME_PATTERN = "yyyy-MM-dd HH:mm";
                CLONG_LDATETIME_PATTERN = "yyyy-MMM-dd HH:mm";
                CLONG_LDATETIMESS_PATTERN = "yyyy-MM-dd HH:mm:ss";
                CNULL_DATE_PATTERN = "yyyy/MM/dd";
                CNULL_SHORTDATE_PATTERN = "yy/MM/dd";
                CNULL_DATE = "1899/12/31";
                CNULL_DATE1 = "1899-12-31";
            }
            else if (g_strDateFormat == "MM/dd/yyyy")
            {
                CSHORT_DATE_PATTERN = "MM-dd-yy";
                CLONG_DATE_PATTERN = "MMM-dd-yy";
                CSHORT_LDATE_PATTERN = "MM-dd-yyyy";
                CSHORT_LDATE_PATTERNMask = "00-00-0000";
                CLONG_LDATE_PATTERN = "MMM-dd-yyyy";
                CSHORT_DATETIME_PATTERN = "MM-dd-yy HH:mm";
                CLONG_DATETIME_PATTERN = "MMM-dd-yy HH:mm";
                CSHORT_LDATETIME_PATTERN = "MM-dd-yyyy HH:mm";
                CLONG_LDATETIME_PATTERN = "MMM-dd-yyyy HH:mm";
                CLONG_LDATETIMESS_PATTERN = "MMM-dd-yyyy HH:mm:ss";
                CNULL_DATE_PATTERN = "MM/dd/yyyy";
                CNULL_SHORTDATE_PATTERN = "MM/dd/yy";
                CNULL_DATE = "12/31/1899";
                CNULL_DATE1 = "12-31-1899";
            }
            WanRoadControl.PublicVariable.g_WebUrl = g_WebUrl;
            WanRoadControl.PublicVariable.g_strSDP = g_strDateFormat;
            WanRoadControl.PublicVariable.g_blnEscDone = false;
            WanRoadControl.PublicVariable.CNULL_DATE = CNULL_DATE;
            WanRoadControl.PublicVariable.CLONG_LDATETIMESS_PATTERN = CLONG_LDATETIMESS_PATTERN;
        }

        private string[] ReturnStringListByCSV(string strLine)
        {
            //    string[] ListByCSV=new string[] { ""};
            List<string> ListByCSV = new List<string>();
            string[] strList = strLine.Split(',');
            for (int intI = 0; intI < strList.Length; intI++)
            {
                ListByCSV.Add(strList[intI]);
                if (strList[intI].Substring(0, 1) == "\"")
                {
                    ListByCSV[ListByCSV.Count - 1] = strList[intI].Substring(1, strList[intI].Length - 1);
                    int intJ = intI + 1;
                    while (strList[intJ].Substring(strList[intJ].Length - 1, 1) != "\"")
                    {
                        ListByCSV[ListByCSV.Count - 1] = ListByCSV[ListByCSV.Count - 1] + "," + strList[intJ];
                        intJ = intJ + 1;
                    }
                    ListByCSV[ListByCSV.Count - 1] = ListByCSV[ListByCSV.Count - 1] + "," + strList[intJ].Substring(0, strList[intJ].Length - 1);
                    if (intI == intJ)
                    { ListByCSV[ListByCSV.Count - 1] = strList[intI].Substring(1, strList[intI].Length - 2); }
                    intI = intJ;
                }
            }
            return ListByCSV.ToArray();
        }

        DataTable ReadDataTableFromExcel(string filepath)
        {
            if (filepath.Trim() == "") { return null; }
            string strFileUpper = filepath.Trim().ToUpper();
            if (strFileUpper.Substring(strFileUpper.Length - ".XLS".Length - 1, ".XLS".Length) == ".XLS" || strFileUpper.Substring(strFileUpper.Length - ".XLS".Length - 1, ".XLSX".Length) == ".XLSX")
            {
                OleDbConnection MyOleDbCn = new OleDbConnection();
                OleDbCommand MyOleDbCmd = new OleDbCommand();
                MyOleDbCn.ConnectionString = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + @filepath + @";Extended ProPerties=""Excel 8.0;HDR=Yes;""";
                MyOleDbCn.Open();
                MyOleDbCmd.Connection = MyOleDbCn;
                MyOleDbCmd.CommandType = CommandType.Text;
                DataTable dt = MyOleDbCn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new Object[] { });
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int intI = 0; intI < dt.Columns.Count; intI++)
                    {
                        if (CheckNull(dt.Rows[0][intI]) != "")
                        {
                            OleDbDataAdapter oda = new OleDbDataAdapter("select * from [" + CheckNull(dt.Rows[0][intI]) + "]", MyOleDbCn);
                            DataTable ds = new System.Data.DataTable();
                            try
                            { oda.Fill(ds); }
                            catch (Exception ex)
                            { ex.Data.Clear(); }
                            MyOleDbCn.Close();
                            return ds;
                        }
                    }
                }
                MyOleDbCn.Close();
            }
            else if (strFileUpper.Substring(strFileUpper.Length - ".CSV".Length - 1, ".CSV".Length) == ".CSV")
            {

                DataTable dt = new DataTable();
                FileStream fs = new FileStream(filepath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
                string strLine = "";
                string[] aryLine;
                int columnCount = 0;
                Boolean IsFirst = true;
                strLine = sr.ReadLine();
                while (strLine.Trim() != "")
                {
                    aryLine = ReturnStringListByCSV(strLine);
                }
                while (strLine != "")
                {
                    aryLine = ReturnStringListByCSV(strLine);
                    if (IsFirst == true)
                    {
                        if (aryLine[0] != "LAYOUT TYPE")
                        {
                            IsFirst = false;
                            columnCount = aryLine.Length;
                            for (int i = 0; i < columnCount; i++)
                            {
                                dt.Columns.Add((i + 1).ToString());
                            }
                        }
                    }
                    if (IsFirst == true)
                    {
                        IsFirst = false;
                        columnCount = aryLine.Length;
                        for (int intI = 0; intI < columnCount; intI++)
                        {
                            DataColumn dc = new DataColumn(aryLine[intI]);
                            dt.Columns.Add(dc);
                        }
                    }
                    else
                    {
                        DataRow dr = dt.NewRow();
                        for (int intI = 0; intI < columnCount; intI++)
                        { dr[intI] = aryLine[intI]; }
                        dt.Rows.Add(dr);
                    }
                    strLine = sr.ReadLine();
                }
                sr.Close();
                fs.Close();
                return dt;
            }
            return null;
        }

        private int InsertTableRecordByDatatableNew(string strtableName, DataTable dt, Boolean blnReturnTrxNo)
        {
            for (int intI = 0; intI < dt.Rows.Count; intI++)
            {
            string strFieldList = "";
                string strValueList = "";
            if (g_strUserID == "") { g_strUserID = g_strConnUserID; }
            for(int intCol = 0;intCol < dt.Columns .Count;intCol ++)
                {
                    strFieldList = strFieldList + (strFieldList.Trim() == "" ? "" : ",") + dt.Columns[intCol].ColumnName;
                    if (dt.Columns[intCol].ColumnName == "WorkStation")
                    { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(Dns.GetHostName()); }
                    else if (dt.Columns[intCol].ColumnName == "StatusCode")
                    { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull("USE"); }
                    else if (dt.Columns[intCol].ColumnName == "CreateDateTime" || dt.Columns[intCol].ColumnName == "UpdateDateTime")
                    { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + "getdate()"; }
                    else if (dt.Columns[intCol].ColumnName == "CreateBy" || dt.Columns[intCol].ColumnName == "UpdateBy")
                    { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(g_strUserID); }
                    else
                    {
                        if (GetDataType(dt.Columns[intCol].DataType.Name) == 2 && CheckNullDate(dt.Rows[intI][intCol], 2) != CheckNullDate("", 2))
                        {
                            if (CheckNullInt(CheckNullDate(dt.Rows[intI][intCol], 2).ToString( "HHmm"), 1) > 0)
                            { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + "'" + CheckNullDate(dt.Rows[intI][intCol], 2).ToString(CUSEDATABASE_DATETIME_PATTERN) + "'"; }
                            else
                            { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(dt.Rows[intI][intCol], GetDataType(dt.Columns[intCol].DataType.Name)); }                                    
                        }
                        else
                        { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(dt.Rows[intI][intCol], GetDataType(dt.Columns[intCol].DataType.Name)); }                        
                    }
    }
                GetSQLCommandReturnIntNew("Insert into  " + strtableName + " (" + strFieldList + ") Values(" + strValueList + ")");
       }
        if( dt.Rows.Count == 1 && blnReturnTrxNo)
            {int intTrxNo = -1;
            DataTable dtRec = GetSQLCommandReturnDTNew("Select Max(TrxNo) from " + strtableName + " Where WorkStation = " + ReplaceWithNull(Dns.GetHostName()));
            if( dtRec!=null && dtRec.Rows.Count>0)
                {
                intTrxNo =CheckNullInt(dtRec.Rows[0][0], 1);
                return intTrxNo;
              }
            }            
      return -1;
   }
        #endregion

        #region Form Function & DetailEDILogic

        public frmsTrendz()
        {
            InitializeComponent();
            args = System.Environment.GetCommandLineArgs();
            Main();
            setEventHandler();           
        }

        private void frmsTrendz_Load(object sender, EventArgs e)
        {
            m_strYear = DateTime.Now.ToString("yyyy");
            m_strMth = DateTime.Now.ToString("MM");
            m_strDay = DateTime.Now.ToString("dd");
            m_strHour = DateTime.Now.ToString("HH");
            m_strMin = DateTime.Now.ToString("mm");
            m_strSecond = DateTime.Now.ToString("ss");
        if(checkDatabaseLogin())
            {
                getFolderName();

            }           
        this.Close();
        Application.Exit();
        }
   
        private void setEventHandler()
        {
            this.Load += new System.EventHandler(this.frmsTrendz_Load);
        }



        #endregion

        #region Purchase Receipt Export
        //private void
        #endregion

        #region Purchase Receipt Import

        #endregion

        #region Purchase Return Shipment Export

        #endregion

        #region Purchase Returm Shipment Import

        #endregion

        #region Sales Shipment Export

        #endregion

        #region Sales Shipment Import

        #endregion

        #region Sales Return Receipt Export

        #endregion

        #region Sales Return Receipt Import

        #endregion

        #region Transfer Shipment Export

        #endregion

        #region Transfer Shipment Import

        #endregion

        #region Master item 

        #endregion

        #region Item Variant

        #endregion

        #region Item Unit of Measure

        #endregion

        #region Item Barcodes

        #endregion
    }
}
