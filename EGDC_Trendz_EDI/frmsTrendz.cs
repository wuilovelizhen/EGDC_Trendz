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
using System.Text.RegularExpressions;

namespace EGDC_Trendz_EDI
{
    public partial class frmsTrendz : SysFreight.Components.frmMainBase
    {

        #region EDI AND SysFreight CommonFunction ALL
        #region Delcare
        string strLocalConnect = "";
        string m_strExportDatabase = "";
        string m_EDIName = "";
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

        #region SysfreightHelpFunction

        private void AutoSetGridColumnMaxLength(string TableName, WanRoadControl.wrJanusGridEx GridObject)
        {
            DataTable dtGex;
            int intI;
            ArrayList arrayColName;
            ArrayList arrayColLength;
            DataTable dtRec = GetSQLCommandReturnDTNew("Select COLUMN_NAME AS ColName, CHARACTER_MAXIMUM_LENGTH AS ColLength,DATA_TYPE from Information_schema.columns where Table_Name=" + ReplaceWithNull(TableName) + "");
            dtGex = (DataTable)GridObject.DataSource;
            if (dtRec != null && dtGex != null && dtRec.Rows.Count > 0 && dtGex.Columns.Count > 0)
            {
                arrayColName = new ArrayList();
                arrayColLength = new ArrayList();
                for (intI = 0; intI < dtRec.Rows.Count; intI++)
                {
                    arrayColName.Add(CheckNull(dtRec.Rows[intI]["ColName"]));
                    arrayColLength.Add(CheckNullInt(dtRec.Rows[intI]["ColLength"], 1));
                }
                for (intI = 0; intI < dtGex.Columns.Count; intI++)
                {
                    GridObject.RootTable.Columns[dtGex.Columns[intI].ColumnName].MaxLength = CheckNullInt(arrayColLength[arrayColName.IndexOf(dtGex.Columns[intI].ColumnName)], 1);
                }
            }
        }

        private string CheckNull(object ojb)
        {
            short intDateType = 0;
            return Modfunction.CheckNull(ojb, ref intDateType).ToString();
        }

        private object CheckNull(object ojb, short intDateType)
        {
            return Modfunction.CheckNull(ojb, ref intDateType);
        }

        private void setEventHandler()
        {
            this.Load += new System.EventHandler(this.frmsEdiForm_Load);
        }

        public frmsTrendz()
        {
            g_strActiveLanguageType = "en-US";
            g_strDateFormat = "dd/MM/yyyy";
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
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            InitializeComponent();
            args = System.Environment.GetCommandLineArgs();
            Main();
            setEventHandler();
        }

        private void frmsEdiForm_Load(object sender, EventArgs e)
        {
            m_strYear = DateTime.Now.ToString("yyyy");
            m_strMth = DateTime.Now.ToString("MM");
            m_strDay = DateTime.Now.ToString("dd");
            m_strHour = DateTime.Now.ToString("HH");
            m_strMin = DateTime.Now.ToString("mm");
            m_strSecond = DateTime.Now.ToString("ss");
            if (m_EDIName == "")
            {
                Main();
                if (m_EDIName == "")
                {
                    MessageBox.Show("Please set the EDI Name in setting.ini");
                    this.Close();
                    Application.Exit();
                    return;
                }
            }
            if (checkDatabaseLogin())
            {
                getFolderName();
                if (int_SaedTrxNo == "")
                {
                    MessageBox.Show("Please seting in EDI table for EDI name = '" + this.m_EDIName + "'");
                    this.Close();
                    Application.Exit();
                    return;
                }
                EdiDetailFunction();
            }
            this.Close();
            Application.Exit();
        }

        private DateTime CheckNullDate(object ojb, int intDateType)
        {
            short intType = 2;
            return Convert.ToDateTime(Modfunction.CheckNull(ojb, ref intType));
        }

        private Double CheckNullDouble(object ojb, int intDateType)
        {
            short intType = 1;
            return Convert.ToDouble(Modfunction.CheckNull(ojb, ref intType));
        }

        private Double CheckNullDouble(object ojb)
        {
            short intType = 1;
            return Convert.ToDouble(Modfunction.CheckNull(ojb, ref intType));
        }

        private Int32 CheckNullInt(object ojb, int intDateType)
        {
            short intType = 1;
            return Convert.ToInt32(Modfunction.CheckNull(ojb, ref intType));
        }

        private Int32 CheckNullInt(object ojb)
        {
            short intType = 1;
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
            if (intDateType == 2) { intType = 2; }
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
            DataTable dtRec = GetSQLCommandReturnDTNew("Select * from Saed1 Where EdiName = " + ReplaceWithNull(this.m_EDIName));
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                m_strFolder = CheckNull(dtRec.Rows[0]["InFolder"]);
                m_strOUTFolder = CheckNull(dtRec.Rows[0]["OutFolder"]);
                m_LogPath = CheckNull(dtRec.Rows[0]["LogFolder"]);
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
                    intLineItemNo = CheckNullInt(dtRec.Rows[0][0], 1) + 1;
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
                if (System.IO.Path.GetExtension(FI.Name) == strFileType)
                {
                    list.Add(FI.Name);
                }
            }
            return list.ToArray();
        }
        #endregion

        static bool IsMoth(string value)
        {
            switch (value)
            {
                case "Atlas Moth":
                case "Beet Armyworm":
                case "Indian Meal Moth":
                case "Ash Pug":
                case "Latticed Heath":
                case "Ribald Wave":
                case "The Streak":
                    return true;
                default:
                    return false;
            }
        }

        #region CheckAndLoginFunction

        private void saveLog(string strValue)
        {
            string strErrorFieldName = "";
            if (strValue != "") { strErrorFieldName = "_Error"; }
            if (m_LogPath.Trim() == "") { m_LogPath = @Directory.GetCurrentDirectory().Trim() + @"\Log"; }
            if (!Directory.Exists(@m_LogPath)) { Directory.CreateDirectory(@m_LogPath); }
            //if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
            //if (!Directory.Exists(@m_BackupLogPath)) { Directory.CreateDirectory(@m_BackupLogPath); }
            string strWriteLog = "";
            string strRowDelimiter = "\r\n";
            strWriteLog = "";
            strWriteLog = strWriteLog + "Run Date And Time are " + m_strDay + @"\" + m_strMth + @"\" + m_strYear + " " + m_strHour + @":" + m_strMin + @":" + m_strSecond + strRowDelimiter;
            strWriteLog = strWriteLog + strValue + strRowDelimiter;
            string strFileName = "RunProject" + strErrorFieldName + "_" + m_strYear.Substring(2, 2) + m_strMth + m_strDay + m_strHour + m_strMin + m_strSecond;
            StreamWriter wBackupLog;
            StreamWriter wLog;
            wLog = File.CreateText(@m_LogPath + @"\" + @strFileName + @".LOG");
            if (!Directory.Exists(@m_LogPath + @"\Log_Backup")) { Directory.CreateDirectory(@m_LogPath + @"\Log_Backup"); }
            wBackupLog = File.CreateText(@m_LogPath + @"\Log_Backup\" + @strFileName + @".LOG");
            wLog.Write(strWriteLog);
            wBackupLog.Write(strWriteLog);
            wLog.Close();
            wBackupLog.Close();
        }
        private Boolean checkDatabaseLogin()
        {
            DataTable dtRec = GetSQLCommandReturnDTNew("Select 1");
            if (dtRec == null || dtRec.Rows.Count == 0)
            {
                if (m_LogPath.Trim() == "") { m_LogPath = @Directory.GetCurrentDirectory().Trim() + @"\Log"; }
                if (!Directory.Exists(@m_LogPath)) { Directory.CreateDirectory(@m_LogPath); }
                //if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                //if (!Directory.Exists(@m_BackupLogPath)) { Directory.CreateDirectory(@m_BackupLogPath); }
                string strWriteLog = "";
                string strRowDelimiter = "\r\n";
                strWriteLog = "";
                strWriteLog = strWriteLog + "Run Date And Time are " + m_strDay + @"\" + m_strMth + @"\" + m_strYear + " " + m_strHour + @":" + m_strMin + @":" + m_strSecond + strRowDelimiter;
                strWriteLog = strWriteLog + "Can not connect the database" + strRowDelimiter;
                string strFileName = "RunProject_" + m_strYear.Substring(2, 2) + m_strMth + m_strDay + m_strHour + m_strMin + m_strSecond;
                StreamWriter wBackupLog;
                StreamWriter wLog;
                if (!Directory.Exists(@m_LogPath + @"\Log_Backup")) { Directory.CreateDirectory(@m_LogPath + @"\Log_Backup"); }
                wLog = File.CreateText(@m_LogPath + @"\" + @strFileName + @".LOG");
                wBackupLog = File.CreateText(@m_LogPath + @"\Log_Backup\" + @strFileName + @".LOG");
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
            strThree = "EDIName";
            strFore = "";
            m_EDIName = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            this.Text = m_EDIName;
            Modfunction.strSqlConn = strLocalConnect;
            Modfunction.strSqlConn2 = strLocalConnect;
            SysfreightSetting = "PathSetting";
            strThree = "EDIPath";
            strFore = "";
            m_EDIPath = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            SysfreightSetting = "SysfreightSetting";
            strThree = "UpdateBy";
            strFore = "";
            g_strUserID = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim();
            this.UserID = g_strUserID;
            strThree = "LogPath";
            strFore = "";
            if (m_LogPath == "") { m_LogPath = sGetINI(ref sIniFile, ref SysfreightSetting, ref strThree, ref strFore).Trim(); }
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

        private int GetDataType(String strDataType)
        {
            switch (strDataType.ToLower())
            {
                case "datetime":
                    return 2;
                case "decimal":
                    return 1;
                case "double":
                    return 1;
                case "int":
                    return 1;
                case "smallint":
                    return 1;
                case "tinyint":
                    return 1;
                case "System.Int32":
                    return 1;
                case "System.Int16":
                    return 1;
                case "System.Int64":
                    return 1;
                case "Int32":
                    return 1;
                case "Int16":
                    return 1;
                case "Int64":
                    return 1;
                case "varchar":
                    return 0;
                case "nvarchar":
                    return 0;
                case "char":
                    return 0;
                default:
                    return 0;
            }
        }

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
            if (g_strUserID == null || g_strUserID == "") { g_strUserID = this.UserID; }
            for (int intI = 0; intI < dt.Rows.Count; intI++)
            {
                string strFieldList = "";
                string strValueList = "";
                if (g_strUserID == "") { g_strUserID = g_strConnUserID; }
                for (int intCol = 0; intCol < dt.Columns.Count; intCol++)
                {
                    if (blnReturnTrxNo && dt.Columns[intCol].ColumnName == "TrxNo") { continue; }
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
                        if (GetDataType(dt.Columns[intCol].DataType.Name) == 2)
                        {
                            if (CheckNullDate(dt.Rows[intI][intCol], 2) != CheckNullDate("", 2))
                            {

                                if (CheckNullInt(CheckNullDate(dt.Rows[intI][intCol], 2).ToString("HHmm"), 1) > 0)
                                { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + "'" + CheckNullDate(dt.Rows[intI][intCol], 2).ToString(CUSEDATABASE_DATETIME_PATTERN) + "'"; }
                                else
                                { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(dt.Rows[intI][intCol], GetDataType(dt.Columns[intCol].DataType.Name)); }
                            }
                            else
                            {
                                strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(dt.Rows[intI][intCol], GetDataType(dt.Columns[intCol].DataType.Name));
                            }
                        }
                        else
                        { strValueList = strValueList + (strValueList.Trim() == "" ? "" : ",") + ReplaceWithNull(dt.Rows[intI][intCol], GetDataType(dt.Columns[intCol].DataType.Name)); }
                    }
                }
                GetSQLCommandReturnIntNew("Insert into  " + strtableName + " (" + strFieldList + ") Values(" + strValueList + ")");
            }
            if (dt.Rows.Count == 1 && blnReturnTrxNo)
            {
                int intTrxNo = -1;
                DataTable dtRec = GetSQLCommandReturnDTNew("Select Max(TrxNo) from " + strtableName + " Where WorkStation = " + ReplaceWithNull(Dns.GetHostName()));
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    intTrxNo = CheckNullInt(dtRec.Rows[0][0], 1);
                    return intTrxNo;
                }
            }
            return -1;
        }

        private Boolean saveToField(DataTable dtGex, string strFileName, string strColDelimiter, string strColValueDelimiter, Boolean blnCaption)
        {
            if (dtGex == null || dtGex.Rows.Count == 0) { return false; }
            if (strFileName == "") { return false; }
            if (strColDelimiter == "") { strColDelimiter = ","; }
            int i = 0;
            int j = 0;
            string strValue = "";
            List<string> strCurrLine = new List<string> { };
            if (blnCaption)
            {
                for (i = 0; i < dtGex.Rows.Count; i++)
                {
                    if (strValue == "")
                    {
                        strValue = dtGex.Columns[i].ColumnName;
                    }
                    else
                    {
                        strValue = strValue + strColDelimiter + dtGex.Columns[i].ColumnName;
                    }
                }
                strCurrLine.Add(strValue);
            }
            for (i = 0; i < dtGex.Rows.Count; i++)
            {
                strValue = "";
                for (j = 0; j < dtGex.Columns.Count; j++)
                {
                    if (strValue == "")
                    {
                        if (strColDelimiter == "," && strColValueDelimiter == "")
                        {
                            if (CheckNull(dtGex.Rows[i][j]).IndexOf(',') > 0)
                            {
                                strValue = "\"" + CheckNull(dtGex.Rows[i][j]) + "\"";
                            }
                            else
                            {
                                strValue = CheckNull(dtGex.Rows[i][j]);
                            }
                        }
                        else
                        {
                            strValue = strColValueDelimiter + CheckNull(dtGex.Rows[i][j]) + strColValueDelimiter;
                        }
                    }
                    else
                    {
                        if (strColDelimiter == "," && strColValueDelimiter == "")
                        {
                            if (CheckNull(dtGex.Rows[i][j]).IndexOf(',') > 0)
                            {
                                strValue = strValue + strColDelimiter + "\"" + CheckNull(dtGex.Rows[i][j]) + "\"";
                            }
                            else
                            {
                                strValue = strValue + strColDelimiter + CheckNull(dtGex.Rows[i][j]);
                            }
                        }
                        else
                        {
                            strValue = strValue + strColDelimiter + strColValueDelimiter + CheckNull(dtGex.Rows[i][j]) + strColValueDelimiter;
                        }
                    }
                }
                strCurrLine.Add(strValue);
            }
            //Open Text File
            //string[dtGex.Rows.Count-1] lines/* = { "First line", "Second line", "Third line" }*/;
            // WriteAllLines creates a file, writes a collection of strings to the file,
            // and then closes the file.  You do NOT need to call Flush() or Close().
            if (strCurrLine.Count > 0)
            {
                System.IO.File.WriteAllLines(@strFileName, strCurrLine.ToArray());
                return true;
            }
            return false;
        }

        private string[] getReadFromTXTFile(string strFile)
        {
            List<string> line = new List<string> { };
            if (strFile.Trim() == "") { return line.ToArray(); }
            string[] lines = System.IO.File.ReadAllLines((this.m_strFolder + @"\" + strFile));
            foreach (string lineNew in lines)
            {
                line.Add(lineNew);
            }
            return line.ToArray();
        }

        private string[] getLineDetail(string strLine, char strColumnDelimiter)
        {
            List<string> strValue = new List<string> { };
            string[] strLineNew = strLine.Split(strColumnDelimiter);
            if (strColumnDelimiter == ',')
            {
                for (int intI = 0; intI < strLineNew.Length; intI++)
                {
                    if (strLineNew[intI].Substring(0, 1) == "\"")
                    {
                        int intJ = intI;
                        string strCellText = strLineNew[intI].Substring(1, strLineNew[intI].Length - 1);
                        while (intJ < strLineNew.Length && strLineNew[intJ].Substring(strLineNew[intJ].Length - 1, 1) != "\"")
                        {
                            strCellText = strCellText + "," + strLineNew[intI];
                        }
                        if (intJ < strLineNew.Length)
                        {
                            strCellText = strCellText + "," + strLineNew[intI].Substring(0, strLineNew[intI].Length - 1);
                        }
                        if (intI == intJ)
                        {
                            strCellText = strLineNew[intI].Substring(1, strLineNew[intI].Length - 2);
                        }
                        intI = intJ;
                    }
                    else
                    {
                        strValue.Add(strLineNew[intI]);
                    }
                }
                return strValue.ToArray();
            }
            else
            {
                for (int intI = 0; intI < strLineNew.Length; intI++)
                {
                    if (strLineNew[intI].Substring(0, 1) == "\"" && strLineNew[intI].Substring(strLineNew[intI].Length - 1, 1) == "\"")
                    {
                        strLineNew[intI] = strLineNew[intI].Substring(1, strLineNew[intI].Length - 2);
                    }
                }
                return strLineNew;
            }
        }
        #endregion

        #region Form Function & DetailEDILogic

        string strDefaultWarehouse = "";
        string strDefaultStoreNo = "";
        string strDefaultCustomerCode = "";
        string strDefaultCustomerName = "";
        string strDefaultCustomerAddress1 = "";
        string strDefaultCustomerAddress2 = "";
        string strDefaultCustomerAddress3 = "";
        string strDefaultCustomerAddress4 = "";
        string strDefaultWarehouseName = "";


        private void EdiDetailFunction()
        {
            getDefaultCustomerCode();
            getWarehouseAndStoreNo();
            switch (m_EDIName.ToLower().Trim())
            {
                case "purchase receipt export":
                    UploadPurchaseReceiptExport();
                    break;
                case "purchase receipt import":
                    setPurchaseReceiptImport();
                    break;
                case "purchase return shipment export":
                    UploadPurchaseReturnShipmentExport();
                    break;
                case "purchase returm shipment import":
                    setPurchaseReturmShipmentImport();
                    break;
                case "sales shipment export":
                    UploadSalesShipmentExport();
                    break;
                case "sales shipment import":
                    setSalesShipmentImport();
                    break;
                case "sales return receipt export":
                    UploadSalesReturnReceiptExport();
                    break;
                case "sales return receipt import":
                    setSalesReturnReceiptImport();
                    break;
                case "transfer shipment export":
                    UploadTransferShipmentExport();
                    break;
                case "transfer shipment import":
                    setTransferShipmentImport();
                    break;
                default:
                    break;
            }
        }

        private Boolean CheckUploadData(string strTable, string strColum, string strValue)
        {
            DataTable dtRec = GetSQLCommandReturnDTNew("Select " + strColum + " from " + strTable + " Where " + strColum + " = " + ReplaceWithNull( strValue ));
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        private void getDefaultCustomerCode()
        {
            if (m_strFilter1 != "")
            {
                string[] strCusotmer = m_strFilter1.Split('=');
                if (strCusotmer != null && strCusotmer.Length == 2)
                {
                    if (strCusotmer[0].ToLower().Replace(" ", "").IndexOf("CustomerCode".ToLower()) >= 0)
                    {
                        this.strDefaultCustomerCode = strCusotmer[1].Trim();
                        if (strDefaultCustomerCode.Substring(0, 1) == "'") { strDefaultCustomerCode = strDefaultCustomerCode.Substring(1, strDefaultCustomerCode.Length - 1); }
                        if (strDefaultCustomerCode.Substring(strDefaultCustomerCode.Length - 1, 1) == "'") { strDefaultCustomerCode = strDefaultCustomerCode.Substring(0, strDefaultCustomerCode.Length - 1); }
                    }
                }
            }
        }

        private void getWarehouseAndStoreNo()
        {
            DataTable dtRec = GetSQLCommandReturnDTNew("Select * from Impa1 ");
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                if (this.strDefaultCustomerCode == "") { this.strDefaultCustomerCode = CheckNull(dtRec.Rows[0]["DefaultCustomerCode"]); }
                this.strDefaultStoreNo = CheckNull(dtRec.Rows[0]["DefaultStoreNo"]);
                this.strDefaultWarehouse = CheckNull(dtRec.Rows[0]["DefaultWarehouseCode"]);
            }
        }

        private DataTable setDefaultCustomerNameAddress(DataTable dt)
        {
            if (strDefaultCustomerCode == "") { return dt; }
            if (strDefaultCustomerName == "")
            {
                DataTable dtRec = GetSQLCommandReturnDTNew("select BusinessPartyName,Address1,Address2,Address3,Address4 From rcbp1 Where BusinessPartyCode = " + ReplaceWithNull(strDefaultCustomerCode));
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    if (dt.Columns.Contains("CustomerCode")) { dt.Rows[0]["CustomerCode"] = strDefaultCustomerCode; }
                    if (dt.Columns.Contains("CustomerName")) { dt.Rows[0]["CustomerName"] = dtRec.Rows[0]["BusinessPartyName"]; }
                    if (dt.Columns.Contains("CustomerAddress1")) { dt.Rows[0]["CustomerAddress1"] = dtRec.Rows[0]["Address1"]; }
                    if (dt.Columns.Contains("CustomerAddress2")) { dt.Rows[0]["CustomerAddress2"] = dtRec.Rows[0]["Address2"]; }
                    if (dt.Columns.Contains("CustomerAddress3")) { dt.Rows[0]["CustomerAddress3"] = dtRec.Rows[0]["Address3"]; }
                    if (dt.Columns.Contains("CustomerAddress4")) { dt.Rows[0]["CustomerAddress4"] = dtRec.Rows[0]["Address4"]; }
                    if (dt.Columns.Contains("Address1")) { dt.Rows[0]["Address1"] = dtRec.Rows[0]["Address1"]; }
                    if (dt.Columns.Contains("Address2")) { dt.Rows[0]["Address2"] = dtRec.Rows[0]["Address2"]; }
                    if (dt.Columns.Contains("Address3")) { dt.Rows[0]["Address3"] = dtRec.Rows[0]["Address3"]; }
                    if (dt.Columns.Contains("Address4")) { dt.Rows[0]["Address4"] = dtRec.Rows[0]["Address4"]; }
                    strDefaultCustomerName = CheckNull(dtRec.Rows[0]["BusinessPartyName"]);
                    strDefaultCustomerAddress1 = CheckNull(dtRec.Rows[0]["Address1"]);
                    strDefaultCustomerAddress2 = CheckNull(dtRec.Rows[0]["Address2"]);
                    strDefaultCustomerAddress3 = CheckNull(dtRec.Rows[0]["Address3"]);
                    strDefaultCustomerAddress4 = CheckNull(dtRec.Rows[0]["Address4"]);
                }
                if (strDefaultWarehouse != "")
                {
                    dtRec = GetSQLCommandReturnDTNew("Select WarehouseName from Whwh1 where WarehouseCode = " + ReplaceWithNull(strDefaultWarehouse));
                    if (dtRec != null && dtRec.Rows.Count > 0)
                    {
                        strDefaultWarehouseName = CheckNull(dtRec.Rows[0]["WarehouseName"]);
                        if (dt.Columns.Contains("WarehouseName")) { dt.Rows[0]["WarehouseName"] = strDefaultWarehouseName; }
                    }
                }
            }
            else
            {
                if (dt.Columns.Contains("CustomerCode")) { dt.Rows[0]["CustomerCode"] = strDefaultCustomerCode; }
                if (dt.Columns.Contains("CustomerName")) { dt.Rows[0]["CustomerName"] = strDefaultCustomerName; }
                if (dt.Columns.Contains("CustomerAddress1")) { dt.Rows[0]["CustomerAddress1"] = strDefaultCustomerAddress1; }
                if (dt.Columns.Contains("CustomerAddress2")) { dt.Rows[0]["CustomerAddress2"] = strDefaultCustomerAddress2; }
                if (dt.Columns.Contains("CustomerAddress3")) { dt.Rows[0]["CustomerAddress3"] = strDefaultCustomerAddress3; }
                if (dt.Columns.Contains("CustomerAddress4")) { dt.Rows[0]["CustomerAddress4"] = strDefaultCustomerAddress4; }
                if (dt.Columns.Contains("Address1")) { dt.Rows[0]["Address1"] = strDefaultCustomerAddress1; }
                if (dt.Columns.Contains("Address2")) { dt.Rows[0]["Address2"] = strDefaultCustomerAddress2; }
                if (dt.Columns.Contains("Address3")) { dt.Rows[0]["Address3"] = strDefaultCustomerAddress3; }
                if (dt.Columns.Contains("Address4")) { dt.Rows[0]["Address4"] = strDefaultCustomerAddress4; }
                if (dt.Columns.Contains("WarehouseName")) { dt.Rows[0]["WarehouseName"] = strDefaultWarehouseName; }
            }
            return dt;
        }

        private int getTrxNoSaveGenerateNumber(string strColumn, string strTable, string strPaColumn, DataTable dt)
        {
            var row = dt.Rows[0];
            row[strColumn] = getGoodsReceiptNo(strTable.Substring(0, strTable.Length - 1), strPaColumn, dt);
            return this.InsertTableRecordByDatatableNew(strTable, dt, true);
        }

        private int GetTrxNoSaveDateRemoveTrxNo(string strTable, DataTable dt)
        {
            var row = dt.Rows[0];
            if (dt.Columns.Contains("TrxNo")) { dt.Columns.Remove(dt.Columns["TrxNo"]); }
            return this.InsertTableRecordByDatatableNew(strTable, dt, true);
        }

        private DataTable setDefaultVenderByCustomer(DataTable dt)
        {
            if (strDefaultCustomerCode == "") { return dt; }
            DataTable dtRec = GetSQLCommandReturnDTNew("select VendorCode,VendorName,Address1,Address2,Address3,Address4 From plvn1 Where VendorCode = (Select Top VendorCode from rcbp1 Where BusinessPartyCode = " + ReplaceWithNull(strDefaultCustomerCode) + ")");
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                if (dt.Columns.Contains("CustomerCode")) { dt.Rows[0]["CustomerCode"] = dtRec.Rows[0]["VendorCode"]; ; }
                if (dt.Columns.Contains("CustomerName")) { dt.Rows[0]["CustomerName"] = dtRec.Rows[0]["BusinessPartyName"]; }
                if (dt.Columns.Contains("CustomerAddress1")) { dt.Rows[0]["CustomerAddress1"] = dtRec.Rows[0]["Address1"]; }
                if (dt.Columns.Contains("CustomerAddress2")) { dt.Rows[0]["CustomerAddress2"] = dtRec.Rows[0]["Address2"]; }
                if (dt.Columns.Contains("CustomerAddress3")) { dt.Rows[0]["CustomerAddress3"] = dtRec.Rows[0]["Address3"]; }
                if (dt.Columns.Contains("CustomerAddress4")) { dt.Rows[0]["CustomerAddress4"] = dtRec.Rows[0]["Address4"]; }
                if (dt.Columns.Contains("Address1")) { dt.Rows[0]["Address1"] = dtRec.Rows[0]["Address1"]; }
                if (dt.Columns.Contains("Address2")) { dt.Rows[0]["Address2"] = dtRec.Rows[0]["Address2"]; }
                if (dt.Columns.Contains("Address3")) { dt.Rows[0]["Address3"] = dtRec.Rows[0]["Address3"]; }
                if (dt.Columns.Contains("Address4")) { dt.Rows[0]["Address4"] = dtRec.Rows[0]["Address4"]; }
            }
            return dt;
        }

        private string getGoodsReceiptNo(string strTable, string strPaColumn, DataTable dt)
        {
            string strGoodNo;
            strGoodNo = CreateGoodsReceiptNo(strTable, dt);
            if (strGoodNo == "" && strPaColumn != "")
            {
                DataTable dtRec = GetSQLCommandReturnDTNew("Select " + strPaColumn + " From Impa1");
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    string strNewNextGoodsReceiptNo = CheckNullInt(dtRec.Rows[0][0], 1).ToString();
                    string strNextGoodsReceiptNo = CheckNull(dtRec.Rows[0][0]);
                    strGoodNo = strNextGoodsReceiptNo;
                    string strGoodsReceiptNo = strNextGoodsReceiptNo;
                    strNewNextGoodsReceiptNo = (CheckNullInt(strNewNextGoodsReceiptNo, 1) + 1).ToString();
                    if (strNewNextGoodsReceiptNo.Length < strGoodsReceiptNo.Length)
                    {
                        while (strNewNextGoodsReceiptNo.Length < strGoodsReceiptNo.Length)
                        {
                            strNewNextGoodsReceiptNo = "0" + strNewNextGoodsReceiptNo;
                        }
                    }
                    GetSQLCommandReturnIntNew("Update Impa1 set " + strPaColumn + " = " + ReplaceWithNull(strNewNextGoodsReceiptNo) + "");
                }
            }
            return strGoodNo;
        }

        private string CreateGoodsReceiptNo(string strTable, DataTable dt)
        {
            string strReceiptNo = "";
            string m_strJobSeqNo = "";
            string m_strPullFrom = "";
            string m_strUpdateNextField = "";
            string[] strArr;
            int intI;
            int intYear = Convert.ToInt32(DateTime.Now.ToString("yyyy"));
            int intMonth = Convert.ToInt32(DateTime.Now.ToString("MM"));
            if (dt.Columns.Contains("ReceiptDate"))
            {
                intYear = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["ReceiptDate"], 2).ToString("yyyy"));
                intMonth = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["ReceiptDate"], 2).ToString("MM"));
            }
            if (dt.Columns.Contains("IssueDateTime"))
            {
                intYear = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["IssueDateTime"], 2).ToString("yyyy"));
                intMonth = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["IssueDateTime"], 2).ToString("MM"));
            }
            if (dt.Columns.Contains("OrderDate"))
            {
                intYear = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["OrderDate"], 2).ToString("yyyy"));
                intMonth = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["OrderDate"], 2).ToString("MM"));
            }
            DataTable dtRec2;
            DataTable dtRec = GetSQLCommandReturnDTNew("Select * From Sanm1 Where NumberType = '" + strTable + "'");
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                var dtRecRow = dtRec.Rows[0];
                if (CheckNull(dtRecRow["Cycle"]) == "") { return ""; }
                if (CheckNull(dtRecRow["Cycle"]) == "C")
                {
                    m_strPullFrom = "NextNo";
                    m_strJobSeqNo = CheckNull(dtRecRow["NextNo"]);
                }
                else if (CheckNull(dtRecRow["Cycle"]) == "M")
                {
                    dtRec2 = GetSQLCommandReturnDTNew("Select * From Sanm2 Where TrxNo = " + CheckNull(dtRecRow["TrxNo"]) + " And Year = '" + intYear.ToString() + "'");
                    if (dtRec2 != null && dtRec2.Rows.Count > 0)
                    {
                        if (intMonth >= 1 && intMonth <= 12)
                        {
                            string strMonth = intMonth.ToString();
                            if (strMonth.Length == 1) { strMonth = "0" + strMonth; }
                            m_strPullFrom = "Mth" + strMonth + "NextNo";
                            m_strJobSeqNo = CheckNull(dtRec2.Rows[0][m_strPullFrom]);
                        }

                    }
                }
                else if (CheckNull(dtRecRow["Cycle"]) == "Y")
                {
                    dtRec2 = GetSQLCommandReturnDTNew("Select YearNextNo From Sanm2 Where TrxNo = " + CheckNull(dtRecRow["TrxNo"]) + " And Year = '" + intYear.ToString() + "'");
                    if (dtRec2 != null && dtRec2.Rows.Count > 0)
                    {
                        m_strPullFrom = "YearNextNo";
                        m_strJobSeqNo = CheckNull(dtRec2.Rows[0]["YearNextNo"]);
                    }
                }
                string strGoodsIssueNo = "";
                if (CheckNull(dtRecRow["Prefix"]).Trim() != "")
                {
                    strArr = CheckNull(dtRecRow["Prefix"]).Split(',');
                    for (intI = 0; intI < strArr.Length; intI++)
                    {
                        strGoodsIssueNo = strGoodsIssueNo + ReturnPrefixSuffix(strArr[intI], dt);
                    }
                }
                strGoodsIssueNo = strGoodsIssueNo + m_strJobSeqNo;
                if (CheckNull(dtRecRow["Suffix"]).Trim() != "")
                {
                    strArr = CheckNull(dtRecRow["Suffix"]).Split(',');
                    for (intI = 0; intI < strArr.Length; intI++)
                    {
                        strGoodsIssueNo = strGoodsIssueNo + ReturnPrefixSuffix(strArr[intI], dt);
                    }
                }
                strReceiptNo = strGoodsIssueNo;
                if (strReceiptNo.Length > strGoodsIssueNo.Length)
                {
                    strReceiptNo = strReceiptNo.Substring(strReceiptNo.Length - strGoodsIssueNo.Length, strGoodsIssueNo.Length);
                }
                m_strUpdateNextField = CheckUpdateFieldLength(ref m_strJobSeqNo);
                if (m_strPullFrom != "" && m_strUpdateNextField != "")
                {
                    if (CheckNull(dtRecRow["Cycle"]) == "C")
                    {
                        GetSQLCommandReturnIntNew("Update Sanm1 Set " + m_strPullFrom + " = " + ReplaceWithNull(m_strUpdateNextField) + " Where TrxNo = " + CheckNull(dtRecRow["TrxNo"]));
                    }
                    else
                    {
                        GetSQLCommandReturnIntNew("Update Sanm2 Set " + m_strPullFrom + " = " + ReplaceWithNull(m_strUpdateNextField) + " Where TrxNo = " + CheckNull(dtRecRow["TrxNo"]) + " And Year = '" + intYear.ToString() + "'");
                    }
                }
            }
            return strReceiptNo;
        }

        private string ReturnPrefixSuffix(string strPrefixSuffix, DataTable dt)
        {
            string strReturn = "";
            try
            {
                if (strPrefixSuffix == "MM")
                {
                    if (dt.Columns.Contains("ReceiptDate")) { strReturn = this.CheckNullDate(dt.Rows[0]["ReceiptDate"], 2).ToString("MM"); }
                    if (dt.Columns.Contains("IssueDateTime")) { strReturn = this.CheckNullDate(dt.Rows[0]["IssueDateTime"], 2).ToString("MM"); }
                    if (dt.Columns.Contains("OrderDate")) { strReturn = this.CheckNullDate(dt.Rows[0]["OrderDate"], 2).ToString("MM"); }
                    if (strReturn == "") { strReturn = DateTime.Now.ToString("MM"); }
                }
                if (strPrefixSuffix == "M")
                {
                    int intMth = -1;
                    if (dt.Columns.Contains("ReceiptDate")) { intMth = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["ReceiptDate"], 2).ToString("MM")); }
                    if (dt.Columns.Contains("IssueDateTime")) { intMth = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["IssueDateTime"], 2).ToString("MM")); }
                    if (dt.Columns.Contains("OrderDate")) { intMth = Convert.ToInt32(this.CheckNullDate(dt.Rows[0]["OrderDate"], 2).ToString("MM")); }
                    if (intMth == -1) { intMth = Convert.ToInt32(DateTime.Now.ToString("MM")); }
                    if (intMth == 12) { strReturn = "D"; }
                    if (intMth == 11) { strReturn = "N"; }
                    if (intMth == 10) { strReturn = "O"; }
                    if (strReturn == "") { strReturn = intMth.ToString(); }

                }
                if (strPrefixSuffix == "YYYY" || strPrefixSuffix == "YY")
                {
                    if (dt.Columns.Contains("ReceiptDate")) { strReturn = this.CheckNullDate(dt.Rows[0]["ReceiptDate"], 2).ToString(strPrefixSuffix.ToLower()); }
                    if (dt.Columns.Contains("IssueDateTime")) { strReturn = this.CheckNullDate(dt.Rows[0]["IssueDateTime"], 2).ToString(strPrefixSuffix.ToLower()); }
                    if (dt.Columns.Contains("OrderDate")) { strReturn = this.CheckNullDate(dt.Rows[0]["OrderDate"], 2).ToString(strPrefixSuffix.ToLower()); }
                    if (strReturn == "") { strReturn = DateTime.Now.ToString(strPrefixSuffix.ToLower()); }
                }
                if (strPrefixSuffix == "Y")
                {
                    if (dt.Columns.Contains("ReceiptDate")) { strReturn = this.CheckNullDate(dt.Rows[0]["ReceiptDate"], 2).ToString("yy"); }
                    if (dt.Columns.Contains("IssueDateTime")) { strReturn = this.CheckNullDate(dt.Rows[0]["IssueDateTime"], 2).ToString("yy"); }
                    if (dt.Columns.Contains("OrderDate")) { strReturn = this.CheckNullDate(dt.Rows[0]["OrderDate"], 2).ToString("yy"); }
                    if (strReturn == "") { strReturn = DateTime.Now.ToString("yy"); }
                    if (strReturn != "") { strReturn = strReturn.Substring(1, 1); }
                }
                if (strPrefixSuffix == "NN")
                {
                    strReturn = "00";
                }
                if (strPrefixSuffix == "N")
                {
                    strReturn = "0";
                }
                if (strPrefixSuffix == "CUST")
                {
                    if (dt.Columns.Contains("CustomerCode")) { strPrefixSuffix = CheckNull(dt.Rows[0]["CustomerCode"]); }
                }
                if (strPrefixSuffix.Substring(0, 1) == "F")
                {
                    strReturn = strPrefixSuffix.Substring(1, strPrefixSuffix.Length - 1);
                }
            }
            catch (Exception ex)
            {
                ex.Data.Clear();
            }
            return strReturn;
        }
        #endregion

        #region Purchase Receipt Export

        private Boolean getCheckErrorProductCode(string strProductCode)
        {
            DataTable dtRec = GetSQLCommandReturnDTNew("Select TrxNo From Impr1 Where ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(strDefaultCustomerCode));
            if (dtRec == null || dtRec.Rows.Count == 0) { return true; }
            return false;
        }

        private DataTable UpdateImgrTotal(DataTable dtImgr1, DataTable dtImgr2)
        {
            if (dtImgr2 != null && dtImgr2.Rows.Count > 0)
            {
                for (int intI = 0; intI < dtImgr2.Rows.Count; intI++)
                {
                    dtImgr1.Rows[0]["TotalReceiptQty"] = CheckNullInt(dtImgr1.Rows[0]["TotalReceiptQty"]) + CheckNullInt(dtImgr2.Rows[intI]["WholeQty"]);
                    dtImgr1.Rows[0]["TotalReceiptQty1"] = CheckNullInt(dtImgr1.Rows[0]["TotalReceiptQty1"]) + CheckNullInt(dtImgr2.Rows[intI]["LooseQty"]);
                    dtImgr1.Rows[0]["TotalReceiptSpace"] = CheckNullDouble(dtImgr1.Rows[0]["TotalReceiptSpace"]) + CheckNullDouble(dtImgr2.Rows[intI]["SpaceArea"]);
                    dtImgr1.Rows[0]["TotalReceiptVolume"] = CheckNullDouble(dtImgr1.Rows[0]["TotalReceiptVolume"]) + CheckNullDouble(dtImgr2.Rows[intI]["Volume"]);
                    dtImgr1.Rows[0]["TotalReceiptWeight"] = CheckNullDouble(dtImgr1.Rows[0]["TotalReceiptWeight"]) + CheckNullDouble(dtImgr2.Rows[intI]["Weight"]);
                }
                dtImgr1.Rows[0]["TotalItem"] = dtImgr2.Rows.Count;
            }
            return dtImgr1;
        }

        private DataTable UpdateImgiTotal(DataTable dtImgi1, DataTable dtImgi2)
        {
            if (dtImgi2 != null && dtImgi2.Rows.Count > 0)
            {
                for (int intI = 0; intI < dtImgi2.Rows.Count; intI++)
                {
                    dtImgi1.Rows[0]["TotalIssueQty"] = CheckNullInt(dtImgi1.Rows[0]["TotalIssueQty"]) + CheckNullInt(dtImgi2.Rows[intI]["WholeQty"]);
                    dtImgi1.Rows[0]["TotalIssueQty1"] = CheckNullInt(dtImgi1.Rows[0]["TotalIssueQty1"]) + CheckNullInt(dtImgi2.Rows[intI]["LooseQty"]);
                    dtImgi1.Rows[0]["TotalIssueSpace"] = CheckNullDouble(dtImgi1.Rows[0]["TotalIssueSpace"]) + CheckNullDouble(dtImgi2.Rows[intI]["SpaceArea"]);
                    dtImgi1.Rows[0]["TotalIssueVolume"] = CheckNullDouble(dtImgi1.Rows[0]["TotalIssueVolume"]) + CheckNullDouble(dtImgi2.Rows[intI]["Volume"]);
                    dtImgi1.Rows[0]["TotalIssueWeight"] = CheckNullDouble(dtImgi1.Rows[0]["TotalIssueWeight"]) + CheckNullDouble(dtImgi2.Rows[intI]["Weight"]);
                }
                dtImgi1.Rows[0]["TotalItem"] = dtImgi2.Rows.Count;
            }
            return dtImgi1;
        }

        private void UploadPurchaseReceiptExport()
        {
            string[] strFileList = this.getFileList(".txt");
            string strErrorProductAll = "";
            string strErrorMessageAll = "";
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImgr1 = null;
                DataTable dtImgr2 = null;
                DataTable dtImpo1 = null;
                DataTable dtImpo2 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    if (dtImgr1 == null)
                    {
                        dtImgr1 = GetSQLCommandReturnDTNew("Select Top 0 * from imgr1");
                        dtImgr2 = GetSQLCommandReturnDTNew("Select Top 0 * from imgr2");
                        dtImpo1 = GetSQLCommandReturnDTNew("Select Top 0 * from impo1");
                        dtImpo2 = GetSQLCommandReturnDTNew("Select Top 0 * from impo2");
                    }
                    string strErrorProduct = "";
                    string strCurrentProduct = "";
                    string strErrorMessage = "";
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        List<string> strPurchaseOrderNoList = new List<string> { };
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 21)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 21") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Purchase Order No", "txt file format error-> Column Count <> 21", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 21" + "|";
                                    continue;
                                }
                            }
                            try
                            {
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                if (this.getCheckErrorProductCode("-" + strValue[12] + "-" + strValue[16]))
                                {
                                    if (strErrorProduct.IndexOf("-" + strValue[12] + "-" + strValue[16]) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strValue[0], "Purchase Order No", strValue[0], "Fail");
                                        strErrorProduct = strErrorProduct + "-" + strValue[12] + "-" + strValue[16] + "|";
                                    }
                                    continue;
                                }                               
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0)
                                {
                                    saveSaed2(strFileList[intI], "", strValue[0], "Purchase Order No", ex.Message, "Fail");
                                    strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|";
                                }
                            }
                        }
                        if (strErrorProduct != "")
                        {
                            strErrorProduct = strErrorProduct.Substring(0, strErrorProduct.Length - 1);
                            strErrorProductAll = strErrorProductAll + "\r\n" + strFileList[intI] + "\r\n Below Product not in sysfreight Product master (Customer : " + this.strDefaultCustomerCode + "), and not upload.\r\n" + strErrorProduct + "\r\n";
                            continue;
                        }
                        else if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + strErrorMessage + "\r\n";
                            continue;
                        }
                        strPurchaseOrderNoList.Clear();
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            try
                            {
                                string[] strValue = this.getLineDetail(strValueList[intJ], '|');                                
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                strCurrentProduct = "-" + strValue[12] + "-" + strValue[16];
                                strPurchaseOrderNoList.Add(strValue[0]);
                                if (CheckUploadData("Imgr1", "PurchaseOrderNo", strValue[0]))
                                {
                                    if (strErrorMessage.IndexOf("Purchase Order No:" + strValue[0] + " already exist") < 0) { strErrorMessage = strErrorMessage + "Purchase Order No:" + strValue[0] + " already exist | "; }
                                    continue;
                                }
                                if (strValue != null && strValue.Length == 21)
                                {
                                    dtImgr1.Rows.Clear();
                                    dtImgr2.Rows.Clear();
                                    dtImpo1.Rows.Clear();
                                    dtImpo2.Rows.Clear();
                                    dtImgr1.Rows.Add(dtImgr1.NewRow());
                                    dtImpo1.Rows.Add(dtImpo1.NewRow());
                                    dtImgr1.Rows[0]["PurchaseOrderNo"] = strValue[0];
                                    dtImpo1.Rows[0]["PurchaseOrderNo"] = strValue[0];
                                    dtImgr1.Rows[0]["ReceiptDate"] = this.CheckNullDate(strValue[3], 2);
                                    dtImpo1.Rows[0]["OrderDate"] = dtImgr1.Rows[0]["ReceiptDate"];
                                    dtImgr1.Rows[0]["ReceiveFrom"] = strValue[6];
                                    dtImgr1.Rows[0]["WarehouseCode"] = this.strDefaultWarehouse;
                                    dtImgr1 = setDefaultCustomerNameAddress(dtImgr1);
                                    dtImpo1 = this.setDefaultVenderByCustomer(dtImpo1);
                                    int intImgrTrxNo = -1;
                                    int intImpoTrxNo = -1;
                                    for (int intDetail = intJ; intDetail < strValueList.Length; intDetail++)
                                    {
                                        string[] strValueNew = this.getLineDetail(strValueList[intDetail], '|');
                                        if (strValueNew != null && strValueNew.Length >= 21)
                                        {
                                            if (strValue[0] == strValueNew[0])
                                            {
                                                dtImgr2.Rows.Add(dtImgr2.NewRow());
                                                dtImpo2.Rows.Add(dtImpo2.NewRow());
                                                dtImgr2 = setPurchaseReceiptExportImgr2(dtImgr2, strValueNew, dtImgr1);
                                                dtImpo2 = setPurchaseReceiptExportImpo2(dtImpo2, strValueNew);
                                            }
                                        }
                                    }
                                    dtImgr1 = UpdateImgrTotal(dtImgr1, dtImgr2);
                                    intImgrTrxNo = this.getTrxNoSaveGenerateNumber("GoodsReceiptNoteNo", "imgr1", "NextGoodsReceiptNo", dtImgr1);
                                    intImpoTrxNo = this.GetTrxNoSaveDateRemoveTrxNo("Impo1", dtImpo1);
                                    if (intImgrTrxNo > 0 && dtImgr2 != null && dtImgr2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dtImgr2.Rows.Count; i++)
                                        {
                                            dtImgr2.Rows[i]["TrxNo"] = intImgrTrxNo;
                                            dtImgr2.Rows[i]["LineItemNo"] = i + 1;
                                        }
                                        this.InsertTableRecordByDatatableNew("Imgr2", dtImgr2, false);
                                        for (int i = 0; i < dtImgr2.Rows.Count; i++)
                                        {
                                            setUpdateImpm1AndImpr1ByImgr(dtImgr2.Rows[i]);
                                            saveSaed2(strFileList[intI], CheckNull(dtImgr1.Rows[0]["GoodsReceiptNoteNo"]), CheckNull(dtImgr1.Rows[0]["PurchaseOrderNo"]), "PurchaseOrderNo", "ProductCode : " + dtImgr2.Rows[i]["ProductCode"] + " <" + getDimensionQty(dtImgr2.Rows[i]) + ">", "Success");
                                        }
                                    }
                                    if (intImpoTrxNo > 0 && dtImpo2 != null && dtImpo2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dtImpo2.Rows.Count; i++)
                                        {
                                            dtImpo2.Rows[i]["TrxNo"] = intImpoTrxNo;
                                            dtImpo2.Rows[i]["LineItemNo"] = i + 1;
                                        }
                                        this.InsertTableRecordByDatatableNew("Impo2", dtImpo2, false);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0) { strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|"; }
                            }
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorProductAll + strErrorMessageAll);
        }

        private DataTable setPurchaseReceiptExportImpo2(DataTable dtImpo2, string[] strValueNew)
        {
            string strProductCode = "-" + strValueNew[12] + "-" + strValueNew[16];
            var dtImpo2Row = dtImpo2.Rows[dtImpo2.Rows.Count - 1];
            dtImpo2Row["UomCode"] = strValueNew[14];
            dtImpo2Row["Qty"] = CheckNullInt(strValueNew[13], 1);
            dtImpo2Row["Description"] = strValueNew[17];
            dtImpo2Row["ProductCode"] = strProductCode;
            dtImpo2Row["PoLineItemNo"] = strValueNew[7];
            dtImpo2Row["Volume"] = CheckNullDouble(strValueNew[20], 1);
            return dtImpo2;
        }

        private DataTable setPurchaseReceiptExportImgr2(DataTable dtImgr2, string[] strValueNew, DataTable dtImgr1)
        {
            DataRow dtImgr2Row = dtImgr2.Rows[dtImgr2.Rows.Count - 1];
            string strProductCode = "-" + strValueNew[12] + "-" + strValueNew[16];
            dtImgr2Row["WarehouseCode"] = strDefaultWarehouse;
            dtImgr2Row["StoreNo"] = strDefaultStoreNo;
            dtImgr2Row["PoNo"] = strValueNew[0];
            dtImgr2Row["UserDefine1"] = strValueNew[0];
            dtImgr2Row["PoLineItemNo"] = strValueNew[7];
            dtImgr2Row["StoreLocation"] = strValueNew[11];
            dtImgr2Row["ProductDescription"] = strValueNew[17];
            dtImgr2Row["ProductCode"] = strProductCode;
            DataTable dtImpr = GetSQLCommandReturnDTNew("Select * from Impr1 Where ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(dtImgr1.Rows[0]["CustomerCode"]));
            dtImgr2Row = this.getImgr2FromImpr(dtImgr2Row, dtImpr, CheckNullInt(strValueNew[13], 1));
            return dtImgr2;
        }

        private string getDimensionQty(DataRow dtRow)
        {
            if (CheckNull(dtRow["DimensionFlag"]) == "1") { return CheckNull(dtRow["PackingQty"]); }
            if (CheckNull(dtRow["DimensionFlag"]) == "2") { return CheckNull(dtRow["WholeQty"]); }
            return CheckNull(dtRow["LooseQty"]);
        }

        #endregion

        #region Purchase Receipt Import

        private void setPurchaseReceiptImport()
        {
            string strFilter = "";
            if (this.m_strFilter1 != "") { strFilter= " AND "+ m_strFilter1; }
            DataTable dtPur = GetSQLCommandReturnDTNew("Select Distinct PurchaseOrderNo from imgr1 Where Imgr1.StatusCode = 'EXE' and (Imgr1.EdiCount = 0 or Imgr1.EdiCount is null) AND (Len(Imgr1.PurchaseOrderNo)>0 OR Imgr1.PurchaseOrderNo <> '') " + strFilter);
            if (dtPur == null || dtPur.Rows.Count == 0) { return; }
            for (int i = 0; i < dtPur.Rows.Count; i++)
            {
                string strSql = "select Imgr1.PurchaseOrderNo,Imgr2.PoLineItemNo,Imgr2.StoreLocation,Imgr2.ProductCode ,";
                strSql = strSql + "case Imgr2.DimensionFlag when '1' then imgr2.PackingQty when '2' then imgr2.WholeQty else imgr2.LooseQty end AS DimensionQty,";
                strSql = strSql + "case Imgr2.DimensionFlag when '1' then(Select Impr1.PackingUomCode from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "when '2' then(Select Impr1.WholeUomCode from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "else  (Select Impr1.LooseUomCode from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) end AS UnitOfMeasureCode,  ";
                strSql = strSql + "case Imgr2.DimensionFlag when '1' then(Select Impr1.PackingPackageSize * Impr1.WholePackageSize from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "when '2' then(Select Impr1.WholePackageSize from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "else  1 end AS QtyPerUnitOfMeasure, ";
                strSql = strSql + "(Select Impo2.Qty From impo2 where Impo2.TrxNo = (Select top 1  TrxNo from impo1 Where Impo1.PurchaseOrderNo = imgr1.PurchaseOrderNo ) AND PoLineItemNo = Imgr2.PoLineItemNo ) AS QtyToReceive, ";
                strSql = strSql + "(Select Impo2.Volume From impo2 where Impo2.TrxNo = (Select top 1  TrxNo from impo1 Where Impo1.PurchaseOrderNo = imgr1.PurchaseOrderNo ) AND PoLineItemNo = Imgr2.PoLineItemNo) AS QtyPerUOM,Imgr1.GoodsReceiptNoteNo ";
                strSql = strSql + " from imgr1 Join imgr2 on Imgr1.TrxNo = imgr2.TrxNo Where Imgr1.StatusCode = 'EXE' and Imgr1.PurchaseOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]);
                DataTable dtRec = GetSQLCommandReturnDTNew(strSql);
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow < dtRec.Rows.Count; intRow++)
                    {
                        string[] strProduct = CheckNull(dtRec.Rows[intRow]["ProductCode"]).Split('-');
                        if (strProduct.Length == 3) { dtRec.Rows[intRow]["ProductCode"] = strProduct[1]; }
                        saveSaed2("", CheckNull(dtRec.Rows[intRow]["GoodsReceiptNoteNo"]), CheckNull(dtRec.Rows[intRow]["PurchaseOrderNo"]), "PurchaseOrderNo", "ProductCode : " + dtRec.Rows[intRow]["ProductCode"] + " <" + dtRec.Rows[intRow]["DimensionQty"] + ">", "Success");
                    }
                    if (dtRec.Columns.Contains("GoodsReceiptNoteNo")) { dtRec.Columns.Remove(dtRec.Columns["GoodsReceiptNoteNo"]); }
                    string strFileName = this.m_strOUTFolder + @"\PurchRcpt-" + CheckNull(dtPur.Rows[i][0]) + "-" + this.m_strYear + this.m_strMth + this.m_strDay + "-" + this.m_strHour + this.m_strMin + @".txt";
                    if (!Directory.Exists(@m_strOUTFolder)) { Directory.CreateDirectory(@m_strOUTFolder); }
                    this.saveToField(dtRec, @m_strOUTFolder, "|", "\"", false);
                    GetSQLCommandReturnIntNew("Update Imgr1 Set EdiCount = isnull(EdiCount,0)+1 Where Imgr1.StatusCode = 'EXE' and Imgr1.PurchaseOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]));
                }
            }
            saveLog("");
        }

        #endregion

        #region Purchase Return Shipment Export

        private DataTable setPurchaseReturnShipmentExportImso2(DataTable dtImso2, string[] strValueNew)
        {
            string strProductCode = "-" + strValueNew[12] + "-" + strValueNew[16];
            var dtImso2Row = dtImso2.Rows[dtImso2.Rows.Count - 1];
            dtImso2Row["UomCode"] = strValueNew[14];
            dtImso2Row["Description"] = strValueNew[17];
            dtImso2Row["ProductCode"] = strProductCode;
            dtImso2Row["SoLineItemNo"] = strValueNew[7];
            dtImso2Row["Volume"] = CheckNullDouble(strValueNew[20], 1);
            return dtImso2;
        }

        private DataTable setPurchaseReturnShipmentExportImgi2(DataTable dtImgi2, string[] strValueNew, DataTable dtImgi1)
        {
            string strProductCode = "-" + strValueNew[12] + "-" + strValueNew[16];
            DataTable dtImpr = GetSQLCommandReturnDTNew("Select * from Impr1 Where ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(dtImgi1.Rows[0]["CustomerCode"]));
            if (dtImpr == null || dtImpr.Rows.Count == 0) { return dtImgi2; }
            int intQty = CheckNullInt(strValueNew[13], 1);
            if (intQty <= 0) { return dtImgi2; }
            DataRow dtImprRows = dtImpr.Rows[0];
            string strSql = "";
            //"FIFO,FIFO|EXPIRY DATE,EXPIRY DATE|MANUFACTURING DATE,MANUFACTURING DATE";
            string[] strlist = getImpmSqlByDim(dtImprRows, dtImgi1);
            strSql = strlist[0];
            string strDimColumnName = strlist[1];
            while (intQty > 0)
            {
                DataTable dtImpm = GetSQLCommandReturnDTNew(strSql);
                if (dtImpm == null || dtImpm.Rows.Count == 0) { return dtImgi2; }
                int intDimQty = CheckNullInt(dtImpm.Rows[0][strDimColumnName], 1);
                if (intDimQty > intQty) { intDimQty = intQty; }
                DataRow dtImpmRows = dtImpm.Rows[0];
                dtImgi2 = PullProductByBatchTrxNo(CheckNullInt(dtImpm.Rows[0]["TrxNo"], 1), dtImgi2, dtImpmRows);
                DataRow dtImgi2Row = dtImgi2.Rows[dtImgi2.Rows.Count - 1];
                dtImgi2Row["SoNo"] = strValueNew[0];
                dtImgi2Row["SoLineItemNo"] = strValueNew[7];
                dtImgi2Row["StoreNo"] = strValueNew[11];
                dtImgi2Row["ProductDescription"] = strValueNew[17];
                dtImgi2Row["ProductCode"] = strProductCode;
                dtImgi2Row = SetFromImpmForImgi(dtImgi2Row, dtImgi2Row, dtImprRows, dtImpr, intDimQty);
                intQty = intQty - intDimQty;
            }


            return dtImgi2;
        }

        private void setImgi2Volumn(string strDimColumn, int lngQty, DataRow dtImpmRow, ref DataRow dtImgiRow)
        {

            Double dblArea = 0.0;
            Double dblVolume = 0.0;
            Double dblUnitVol = 0.0;
            Double dblUnitWt = 0.0;
            Double dblWeight = 0.0;
            int lngBalQty = 0;
            lngBalQty = CheckNullInt(dtImpmRow[strDimColumn], 1);
            dblVolume = CheckNullInt(dtImpmRow["BalanceVolume"], 1);
            dblWeight = CheckNullInt(dtImpmRow["BalanceWeight"], 1);
            dblArea = CheckNullInt(dtImpmRow["BalanceSpaceArea"], 1);
            if (lngQty == lngBalQty)
            {
                dtImgiRow["Volume"] = dblVolume;
                dtImgiRow["Weight"] = dblWeight;
                dtImgiRow["SpaceArea"] = dblArea;
                return;
            }

            dblUnitVol = CheckNullDouble(dtImpmRow["UnitVol"], 1);
            dblUnitWt = CheckNullDouble(dtImpmRow["UnitWt"], 1);
            if (dblUnitVol <= 0 || dblUnitWt <= 0)
            {
                if (CheckNull(dtImpmRow["DimensionFlag"]) == "1")
                {
                    if (dblUnitVol == 0 && CheckNullInt(dtImpmRow["PackingQty"], 1) != 0)
                    {
                        dblUnitVol = CheckNullDouble(dtImpmRow["Volume"], 1) / CheckNullInt(dtImpmRow["PackingQty"], 1);
                    }
                    if (dblUnitWt == 0 && CheckNullInt(dtImpmRow["PackingQty"], 1) != 0)
                    {
                        dblUnitWt = CheckNullDouble(dtImpmRow["Weight"], 1) / CheckNullInt(dtImpmRow["PackingQty"], 1);
                    }
                }
                else if (CheckNull(dtImpmRow["DimensionFlag"]) == "2")
                {
                    if (dblUnitVol == 0 && CheckNullInt(dtImpmRow["WholeQty"], 1) != 0)
                    {
                        dblUnitVol = CheckNullDouble(dtImpmRow["Volume"], 1) / CheckNullInt(dtImpmRow["WholeQty"], 1);
                    }
                    if (dblUnitWt == 0 && CheckNullInt(dtImpmRow["WholeQty"], 1) != 0)
                    {
                        dblUnitWt = CheckNullDouble(dtImpmRow["Weight"], 1) / CheckNullInt(dtImpmRow["WholeQty"], 1);
                    }
                }
                else
                {
                    if (dblUnitVol == 0 && CheckNullInt(dtImpmRow["LooseQty"], 1) != 0)
                    {
                        dblUnitVol = CheckNullDouble(dtImpmRow["Volume"], 1) / CheckNullInt(dtImpmRow["LooseQty"], 1);
                    }
                    if (dblUnitWt == 0 && CheckNullInt(dtImpmRow["LooseQty"], 1) != 0)
                    {
                        dblUnitWt = CheckNullDouble(dtImpmRow["Weight"], 1) / CheckNullInt(dtImpmRow["LooseQty"], 1);
                    }
                }
            }
            if (dblUnitVol > 0)
            {
                dtImgiRow["Volume"] = lngQty * dblUnitVol;
            }
            if (dblUnitWt > 0)
            {
                dtImgiRow["Weight"] = lngQty * dblUnitWt;
            }
            dtImgiRow["SpaceArea"] = lngQty * CheckNullDouble(dtImgiRow["Length"], 1) * CheckNullDouble(dtImgiRow["Width"], 1);
        }

        private void UploadPurchaseReturnShipmentExport()
        {
            string[] strFileList = this.getFileList(".txt");
            string strErrorProductAll = "";
            string strErrorMessageAll = "";
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImgi1 = null;
                DataTable dtImgi2 = null;
                DataTable dtImso1 = null;
                DataTable dtImso2 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    if (dtImgi1 == null)
                    {
                        dtImgi1 = GetSQLCommandReturnDTNew("Select Top 0 * from imgi1");
                        dtImgi2 = GetSQLCommandReturnDTNew("Select Top 0 * from imgi2");
                        dtImso1 = GetSQLCommandReturnDTNew("Select Top 0 * from imso1");
                        dtImso2 = GetSQLCommandReturnDTNew("Select Top 0 * from imso2");
                    }
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        string strErrorProduct = "";
                        string strCurrentProduct = "";
                        string strErrorMessage = "";
                        List<string> strPurchaseOrderNoList = new List<string> { };
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 21)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 21") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Sales Order No", "txt file format error-> Column Count <> 21", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 21" + "|";
                                    continue;
                                }
                            }                           
                            try
                            {
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                if (this.getCheckErrorProductCode("-" + strValue[12] + "-" + strValue[16]))
                                {
                                    if (strErrorProduct.IndexOf("-" + strValue[12] + "-" + strValue[16]) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", strValue[0], "Fail");
                                        strErrorProduct = strErrorProduct + "-" + strValue[12] + "-" + strValue[16] + "|";
                                    }
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0)
                                {
                                    saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", ex.Message, "Fail");
                                    strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|";
                                }
                            }
                        }
                        if (strErrorProduct != "")
                        {
                            strErrorProduct = strErrorProduct.Substring(0, strErrorProduct.Length - 1);
                            strErrorProductAll = strErrorProductAll + "\r\n" + strFileList[intI] + "\r\n Below Product not in sysfreight Product master (Customer : " + this.strDefaultCustomerCode + "), and not upload.\r\n" + strErrorProduct + "\r\n";
                            continue;
                        }
                        else if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + strErrorMessage + "\r\n";
                            continue;
                        }
                        strPurchaseOrderNoList.Clear();
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                            strPurchaseOrderNoList.Add(strValue[0]);
                            if (CheckUploadData("Imgi1", "SalesOrderNo", strValue[0]))
                            {
                                if (strErrorMessage.IndexOf("Sales Order No:" + strValue[0] + " already exist") < 0) { strErrorMessage = strErrorMessage + "Sales Order No:" + strValue[0] + " already exist | "; }
                                continue;
                            }
                            if (strValue != null && strValue.Length == 21)
                            {
                                dtImgi1.Rows.Clear();
                                dtImgi2.Rows.Clear();
                                dtImso1.Rows.Clear();
                                dtImso2.Rows.Clear();
                                dtImgi1.Rows.Add(dtImgi1.NewRow());
                                dtImso1.Rows.Add(dtImso1.NewRow());
                                dtImgi1.Rows[0]["SalesOrderNo"] = strValue[0];
                                dtImso1.Rows[0]["SalesOrderNo"] = strValue[0];
                                dtImgi1.Rows[0]["IssueDateTime"] = this.CheckNullDate(strValue[3], 2);
                                dtImso1.Rows[0]["OrderDate"] = dtImgi1.Rows[0]["IssueDateTime"];
                                dtImgi1 = setDefaultCustomerNameAddress(dtImgi1);
                                dtImso1 = this.setDefaultVenderByCustomer(dtImso1);
                                int intImgiTrxNo = -1;
                                int intImsoTrxNo = -1;
                                for (int intDetail = intJ; intDetail < strValueList.Length; intDetail++)
                                {
                                    string[] strValueNew = this.getLineDetail(strValueList[intDetail], '|');
                                    if (strValueNew != null && strValueNew.Length == 21)
                                    {
                                        if (strValue[0] == strValueNew[0])
                                        {
                                            //dtImgi2.Rows.Add(dtImgi2.NewRow());
                                            dtImso2.Rows.Add(dtImso2.NewRow());
                                            dtImgi2 = setPurchaseReturnShipmentExportImgi2(dtImgi2, strValueNew, dtImgi1);
                                            dtImso2 = setPurchaseReturnShipmentExportImso2(dtImso2, strValueNew);
                                        }
                                    }
                                }
                                if (dtImgi2 == null || dtImgi2.Rows.Count == 0)
                                {
                                    strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + "Sales Order No : " + strValue[0] + ", not Dim Qty to Issue." + "\r\n";
                                    saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", strValue[0], "Fail");
                                    continue;
                                }
                                dtImgi1 = UpdateImgiTotal(dtImgi1, dtImgi2);
                                intImgiTrxNo = this.getTrxNoSaveGenerateNumber("GoodsIssueNoteNo", "imgi1", "NextGoodsIssueNo", dtImgi1);
                                intImsoTrxNo = this.GetTrxNoSaveDateRemoveTrxNo("Imso1", dtImso1);
                                if (intImgiTrxNo > 0 && dtImgi2 != null && dtImgi2.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtImgi2.Rows.Count; i++)
                                    {
                                        dtImgi2.Rows[i]["TrxNo"] = intImgiTrxNo;
                                        dtImgi2.Rows[i]["LineItemNo"] = i + 1;
                                    }
                                    this.InsertTableRecordByDatatableNew("Imgi2", dtImgi2, false);
                                    for (int i = 0; i < dtImgi2.Rows.Count; i++)
                                    {
                                        setUpdateImpm1AndImpr1ByImgi(dtImgi2.Rows[i]);
                                        saveSaed2(strFileList[intI], CheckNull(dtImgi1.Rows[0]["GoodsIssueNoteNo"]), CheckNull(dtImgi1.Rows[0]["SalesOrderNo"]), "SalesOrderNo", "ProductCode : " + dtImgi2.Rows[i]["ProductCode"] + " <" + getDimensionQty(dtImgi2.Rows[i]) + ">", "Success");
                                    }

                                }
                                if (intImsoTrxNo > 0 && dtImso2 != null && dtImso2.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtImso2.Rows.Count; i++)
                                    {
                                        dtImso2.Rows[i]["TrxNo"] = intImsoTrxNo;
                                        dtImso2.Rows[i]["LineItemNo"] = i + 1;
                                    }
                                    this.InsertTableRecordByDatatableNew("Imso2", dtImso2, false);
                                }
                            }
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorProductAll + strErrorMessageAll);
        }

        private DataTable PullProductByBatchTrxNo(int intTrxNo, DataTable dtImgi2, DataRow dtImpmRow)
        {
            dtImgi2.Rows.Add(dtImgi2.NewRow());
            DataRow dtImgiRow = dtImgi2.Rows[dtImgi2.Rows.Count - 1];
            dtImgiRow["WarehouseCode"] = dtImpmRow["WarehouseCode"];
            dtImgiRow["StoreNo"] = dtImpmRow["StoreNo"];
            dtImgiRow["BatchNo"] = dtImpmRow["BatchNo"];
            dtImgiRow["ProductTrxNo"] = dtImpmRow["ProductTrxNo"];
            dtImgiRow["ReceiptMovementTrxNo"] = intTrxNo;
            DataTable dtRec = GetSQLCommandReturnDTNew("SELECT PartNoTrxNo, PartNo,Impm1.ProductName FROM Impm1 LEFT JOIN Impn1 ON Impm1.PartNoTrxNo = Impn1.TrxNo WHERE Impm1.TrxNo = " + ReplaceWithNull(intTrxNo, 1));
            if (dtRec != null && dtRec.Rows.Count > 0)
            {
                dtImgiRow["ProductDescription"] = dtRec.Rows[0]["ProductName"];
                dtImgiRow["PartNoTrxNo"] = dtRec.Rows[0]["PartNoTrxNo"];
            }
            dtImgiRow["Length"] = DBNull.Value;
            dtImgiRow["Width"] = DBNull.Value;
            dtImgiRow["Height"] = DBNull.Value;
            if (CheckNull(dtImgiRow["BatchNo"]) != "O /B")
            {
                dtRec = GetSQLCommandReturnDTNew("SELECT a.DimensionFlag AS 'DimFlag', a.Length AS 'IssueLength', a.Width AS 'IssueWidth', a.Height AS 'IssueHeight' FROM Impm1 LEFT JOIN (SELECT GoodsReceiptNoteNo, LineItemNo, DimensionFlag, Height, Length, Width FROM imgr1 left join imgr2 on imgr1.TrxNo = Imgr2.TrxNo) a ON Impm1.BatchNo = a.GoodsReceiptNoteNo AND Impm1.BatchLineItemNo = a.LineItemNo WHERE Impm1.TrxNo = " + ReplaceWithNull(intTrxNo, 1));
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    dtImgiRow["Length"] = dtRec.Rows[0]["IssueLength"];
                    dtImgiRow["Width"] = dtRec.Rows[0]["IssueWidth"];
                    dtImgiRow["Height"] = dtRec.Rows[0]["IssueHeight"];
                }
            }
            dtImgiRow["ProductDescription"] = dtImpmRow["ProductName"];
            dtImgiRow["LotNo"] = dtImpmRow["LotNo"];
            dtImgiRow["Location"] = dtImpmRow["Location"];
            dtImgiRow["UomCode"] = dtImpmRow["UomCode"];
            dtImgiRow["UserDefine1"] = dtImpmRow["UserDefine1"];
            dtImgiRow["UserDefine3"] = dtImpmRow["UserDefine2"];
            dtImgiRow["PackingQty"] = dtImpmRow["BalancePackingQty"];
            dtImgiRow["WholeQty"] = dtImpmRow["BalanceWholeQty"];
            dtImgiRow["LooseQty"] = dtImpmRow["BalanceLooseQty"];
            dtImgiRow["SpaceArea"] = dtImpmRow["BalanceSpaceArea"];
            dtImgiRow["Weight"] = dtImpmRow["BalanceWeight"];
            dtImgiRow["Volume"] = dtImpmRow["BalanceVolume"];
            dtImgiRow["ManufactureDate"] = dtImpmRow["ManufactureDate"];
            dtImgiRow["ExpiryDate"] = dtImpmRow["ExpiryDate"];
            return dtImgi2;
        }

        private void setUpdateImpm1AndImpr1ByImgr(DataRow dtImgrRow)
        {
            String m_SQlCommandText = "UPDATE Impr1 SET PackingIncomingQty = PackingIncomingQty + " + CheckNullInt(dtImgrRow["PackingQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", WholeIncomingQty = WholeIncomingQty + " + CheckNullInt(dtImgrRow["WholeQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", LooseIncomingQty = LooseIncomingQty + " + CheckNullInt(dtImgrRow["LooseQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + " WHERE TrxNo = " + CheckNullInt(dtImgrRow["ProductTrxNo"], 1).ToString();
            GetSQLCommandReturnIntNew(m_SQlCommandText);
            GetSQLCommandReturnIntNew("UPDATE Impm1 SET ProductName = " + ReplaceWithNull(dtImgrRow["ProductDescription"]) + ",UpdateDateTime=getdate(),UpdateBy='" + g_strUserID + "' WHERE TrxNo = " + ReplaceWithNull(dtImgrRow["ProductTrxNo"], 1));
        }

        private void setUpdateImpm1AndImpr1ByImgi(DataRow dtImgiRow)
        {
            string m_SQlCommandText;
            m_SQlCommandText = "UPDATE Impm1 SET BalancePackingQty = BalancePackingQty - " + CheckNullInt(dtImgiRow["PackingQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", BalanceWholeQty = BalanceWholeQty - " + CheckNullInt(dtImgiRow["WholeQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", BalanceLooseQty = BalanceLooseQty - " + CheckNullInt(dtImgiRow["LooseQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", BalanceSpaceArea = BalanceSpaceArea - " + CheckNullInt(dtImgiRow["SpaceArea"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", BalanceVolume = BalanceVolume - " + CheckNullInt(dtImgiRow["Volume"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", BalanceWeight = BalanceWeight - " + CheckNullInt(dtImgiRow["Weight"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", UpdateDateTime=getdate(),UpdateBy='" + g_strUserID + "' WHERE TrxNo = " + CheckNullInt(dtImgiRow["ReceiptMovementTrxNo"], 1).ToString();
            GetSQLCommandReturnIntNew(m_SQlCommandText);
            m_SQlCommandText = "UPDATE Impr1 SET PackingOnOrderQty = PackingOnOrderQty + " + CheckNullInt(dtImgiRow["PackingQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", WholeOnOrderQty = WholeOnOrderQty + " + CheckNullInt(dtImgiRow["WholeQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + ", LooseOnOrderQty = LooseOnOrderQty + " + CheckNullInt(dtImgiRow["LooseQty"], 1).ToString();
            m_SQlCommandText = m_SQlCommandText + " WHERE TrxNo = " + CheckNullInt(dtImgiRow["ProductTrxNo"], 1).ToString();
            GetSQLCommandReturnIntNew(m_SQlCommandText);
        }

        #endregion

        #region Purchase Returm Shipment Import

        private void setPurchaseReturmShipmentImport()
        {
            string strFilter = "";
            if (this.m_strFilter1 != "") { strFilter = " AND " + m_strFilter1; }
            DataTable dtPur = GetSQLCommandReturnDTNew("Select Distinct SalesOrderNo from imgi1 Where imgi1.StatusCode = 'EXE' and (imgi1.EdiCount = 0 or imgi1.EdiCount is null) AND (Len(imgi1.SalesOrderNo)>0 OR imgi1.SalesOrderNo <> '') " + strFilter);
            if (dtPur == null || dtPur.Rows.Count == 0) { return; }
            for (int i = 0; i < dtPur.Rows.Count; i++)
            {
                string strSql = "select Imgi1.SalesOrderNo,Imgi2.SoLineItemNo,Imgi2.StoreNo,Imgi2.ProductCode ,";
                strSql = strSql + "case Imgi2.DimensionFlag when '1' then Imgi2.PackingQty when '2' then Imgi2.WholeQty else Imgi2.LooseQty end AS DimensionQty,";
                strSql = strSql + "case Imgi2.DimensionFlag when '1' then(Select Impr1.PackingUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " when '2' then(Select Impr1.WholeUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " else  (Select Impr1.LooseUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo ) end AS UnitOfMeasureCode,";
                strSql = strSql + "  case Imgi2.DimensionFlag when '1' then(Select Impr1.PackingPackageSize * Impr1.WholePackageSize from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " when '2' then(Select Impr1.WholePackageSize from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " else  1 end AS QtyPerUnitOfMeasure, ";
                strSql = strSql + "(Select Imso2.Qty From Imso2 where Imso2.TrxNo = (Select top 1  TrxNo from Imso1 Where Imso1.SalesOrderNo = Imgi1.SalesOrderNo ) AND SoLineItemNo = Imgi2.SoLineItemNo ) AS QtyToReceive,";
                strSql = strSql + " (Select Imso2.Volume From Imso2 where Imso2.TrxNo = (Select top 1  TrxNo from Imso1 Where Imso1.SalesOrderNo = Imgi1.SalesOrderNo ) AND SoLineItemNo = Imgi2.SoLineItemNo) AS QtyPerUOM,Imgi1.GoodsIssueNoteNo";
                strSql = strSql + "  from Imgi1 Join Imgi2 on Imgi1.TrxNo = Imgi2.TrxNo Where Imgi1.StatusCode = 'EXE' and Imgi1.SalesOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]);
                DataTable dtRec = GetSQLCommandReturnDTNew(strSql);
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow < dtRec.Rows.Count; intRow++)
                    {
                        string[] strProduct = CheckNull(dtRec.Rows[intRow]["ProductCode"]).Split('-');
                        if (strProduct.Length == 3) { dtRec.Rows[intRow]["ProductCode"] = strProduct[1]; }
                        saveSaed2("", CheckNull(dtRec.Rows[intRow]["GoodsIssueNoteNo"]), CheckNull(dtRec.Rows[intRow]["SalesOrderNo"]), "SalesOrderNo", "ProductCode : " + dtRec.Rows[intRow]["ProductCode"] + " <" + dtRec.Rows[intRow]["DimensionQty"] + ">", "Success");
                    }
                    if (dtRec.Columns.Contains("GoodsIssueNoteNo")) { dtRec.Columns.Remove(dtRec.Columns["GoodsIssueNoteNo"]); }
                    string strFileName = this.m_strOUTFolder + @"\ReturnShip-" + CheckNull(dtPur.Rows[i][0]) + "-" + this.m_strYear + this.m_strMth + this.m_strDay + "-" + this.m_strHour + this.m_strMin + @".txt";
                    if (!Directory.Exists(@m_strOUTFolder)) { Directory.CreateDirectory(@m_strOUTFolder); }
                    this.saveToField(dtRec, @m_strOUTFolder, "|", "\"", false);
                    GetSQLCommandReturnIntNew("Update Imgi1 Set EdiCount = isnull(EdiCount,0)+1 Where Imgi1.StatusCode = 'EXE' and Imgi1.SalesOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]));
                }
            }
            saveLog("");
        }

        #endregion

        #region Sales Shipment Export
        private void UploadSalesShipmentExport()
        {
            string[] strFileList = this.getFileList(".txt");
            string strErrorProductAll = "";
            string strErrorMessageAll = "";
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImgi1 = null;
                DataTable dtImgi2 = null;
                DataTable dtImso1 = null;
                DataTable dtImso2 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    if (dtImgi1 == null)
                    {
                        dtImgi1 = GetSQLCommandReturnDTNew("Select Top 0 * from imgi1");
                        dtImgi2 = GetSQLCommandReturnDTNew("Select Top 0 * from imgi2");
                        dtImso1 = GetSQLCommandReturnDTNew("Select Top 0 * from imso1");
                        dtImso2 = GetSQLCommandReturnDTNew("Select Top 0 * from imso2");
                    }
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        string strErrorProduct = "";
                        string strCurrentProduct = "";
                        string strErrorMessage = "";
                        List<string> strPurchaseOrderNoList = new List<string> { };
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 35)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 35") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Sales Order No", "txt file format error-> Column Count <> 35", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 35" + "|";
                                    continue;
                                }
                            }
                            try
                            {
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                if (this.getCheckErrorProductCode("-" + strValue[14] + "-" + strValue[18]))
                                {
                                    if (strErrorProduct.IndexOf("-" + strValue[14] + "-" + strValue[18]) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", strValue[0], "Fail");
                                        strErrorProduct = strErrorProduct + "-" + strValue[14] + "-" + strValue[18] + "|";
                                    }
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0)
                                {
                                    saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", ex.Message, "Fail");
                                    strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|";
                                }
                            }
                        }
                        if (strErrorProduct != "")
                        {
                            strErrorProduct = strErrorProduct.Substring(0, strErrorProduct.Length - 1);
                            strErrorProductAll = strErrorProductAll + "\r\n" + strFileList[intI] + "\r\n Below Product not in sysfreight Product master (Customer : " + this.strDefaultCustomerCode + "), and not upload.\r\n" + strErrorProduct + "\r\n";
                            continue;
                        }
                        else if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + strErrorMessage + "\r\n";
                            continue;
                        }
                        strPurchaseOrderNoList.Clear();
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                            strPurchaseOrderNoList.Add(strValue[0]);
                            if (CheckUploadData("Imgi1", "SalesOrderNo", strValue[0]))
                            {
                                if (strErrorMessage.IndexOf("Sales Order No:" + strValue[0] + " already exist") < 0) { strErrorMessage = strErrorMessage + "Sales Order No:" + strValue[0] + " already exist | "; }
                                continue;
                            }
                            if (strValue != null && strValue.Length == 35)
                            {
                                dtImgi1.Rows.Clear();
                                dtImgi2.Rows.Clear();
                                dtImso1.Rows.Clear();
                                dtImso2.Rows.Clear();
                                dtImgi1.Rows.Add(dtImgi1.NewRow());
                                dtImso1.Rows.Add(dtImso1.NewRow());
                                dtImgi1.Rows[0]["SalesOrderNo"] = strValue[0];
                                dtImso1.Rows[0]["SalesOrderNo"] = strValue[0];
                                dtImgi1.Rows[0]["IssueDateTime"] = this.CheckNullDate(strValue[3], 2);
                                dtImso1.Rows[0]["OrderDate"] = dtImgi1.Rows[0]["IssueDateTime"];
                                dtImso1.Rows[0]["IssueTo"] = strValue[8];
                                dtImso1.Rows[0]["DeliveryToName"] = strValue[26];
                                dtImso1.Rows[0]["DeliveryToAddress1"] = strValue[27];
                                dtImso1.Rows[0]["DeliveryToAddress2"] = strValue[28];
                                dtImso1.Rows[0]["Description1"] = strValue[29].Substring(0, 50);
                                if (strValue[29].Length > 50) { dtImso1.Rows[0]["Description2"] = strValue[29].Substring(50, strValue[29].Length - 50); }
                                dtImso1.Rows[0]["DeliveryToContactName"] = strValue[30];
                                dtImso1.Rows[0]["DeliveryToAddress3"] = strValue[31];
                                dtImso1.Rows[0]["DeliveryToAddress4"] = strValue[32];
                                dtImgi1 = setDefaultCustomerNameAddress(dtImgi1);
                                dtImso1 = this.setDefaultVenderByCustomer(dtImso1);
                                int intImgiTrxNo = -1;
                                int intImsoTrxNo = -1;
                                for (int intDetail = intJ; intDetail < strValueList.Length; intDetail++)
                                {
                                    string[] strValueNew = this.getLineDetail(strValueList[intDetail], '|');
                                    if (strValueNew != null && strValueNew.Length == 35)
                                    {
                                        if (strValue[0] == strValueNew[0])
                                        {
                                            //dtImgi2.Rows.Add(dtImgi2.NewRow());
                                            dtImso2.Rows.Add(dtImso2.NewRow());
                                            dtImgi2 = setSalesShipmentExportImgi2(dtImgi2, strValueNew, dtImgi1);
                                            dtImso2 = setSalesShipmentExportImso2(dtImso2, strValueNew);
                                        }
                                    }
                                }
                                if (dtImgi2 == null || dtImgi2.Rows.Count == 0)
                                {
                                    strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + "Sales Order No : " + strValue[0] + ", not Dim Qty to Issue." + "\r\n";
                                    saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", strValue[0], "Fail");
                                    continue;
                                }
                                dtImgi1 = UpdateImgiTotal(dtImgi1, dtImgi2);
                                intImgiTrxNo = this.getTrxNoSaveGenerateNumber("GoodsIssueNoteNo", "imgi1", "NextGoodsIssueNo", dtImgi1);
                                intImsoTrxNo = this.GetTrxNoSaveDateRemoveTrxNo("Imso1", dtImso1);
                                if (intImgiTrxNo > 0 && dtImgi2 != null && dtImgi2.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtImgi2.Rows.Count; i++)
                                    {
                                        dtImgi2.Rows[i]["TrxNo"] = intImgiTrxNo;
                                        dtImgi2.Rows[i]["LineItemNo"] = i + 1;
                                    }
                                    this.InsertTableRecordByDatatableNew("Imgi2", dtImgi2, false);
                                    for (int i = 0; i < dtImgi2.Rows.Count; i++)
                                    {
                                        setUpdateImpm1AndImpr1ByImgi(dtImgi2.Rows[i]);
                                        saveSaed2(strFileList[intI], CheckNull(dtImgi1.Rows[0]["GoodsIssueNoteNo"]), CheckNull(dtImgi1.Rows[0]["SalesOrderNo"]), "SalesOrderNo", "ProductCode : " + dtImgi2.Rows[i]["ProductCode"] + " <" + getDimensionQty(dtImgi2.Rows[i]) + ">", "Success");
                                    }

                                }
                                if (intImsoTrxNo > 0 && dtImso2 != null && dtImso2.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtImso2.Rows.Count; i++)
                                    {
                                        dtImso2.Rows[i]["TrxNo"] = intImsoTrxNo;
                                        dtImso2.Rows[i]["LineItemNo"] = i + 1;
                                    }
                                    this.InsertTableRecordByDatatableNew("Imso2", dtImso2, false);
                                }
                            }
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorProductAll + strErrorMessageAll);
        }

        private DataTable setSalesShipmentExportImso2(DataTable dtImso2, string[] strValueNew)
        {
            string strProductCode = "-" + strValueNew[14] + "-" + strValueNew[18];
            var dtImso2Row = dtImso2.Rows[dtImso2.Rows.Count - 1];
            dtImso2Row["UomCode"] = strValueNew[16];
            dtImso2Row["Description"] = strValueNew[19];
            dtImso2Row["ProductCode"] = strProductCode;
            dtImso2Row["SoLineItemNo"] = strValueNew[9];
            dtImso2Row["Volume"] = CheckNullDouble(strValueNew[17], 1);
            return dtImso2;
        }

        private string[] getImpmSqlByDim(DataRow dtImprRows, DataTable dtImgi1)
        {
            List<string> list = new List<string> { };
            string strOrderBy = "";
            string strOrderFile = ",'' AS OrderField ";
            if (CheckNull(dtImprRows["IssueMethod"]) == "FIFO" || CheckNull(dtImprRows["IssueMethod"]) == "")
            {
                strOrderBy = " Order By OrderField,[TrxNo] ";
                strOrderFile = ",'' AS OrderField ";
            }
            else if (CheckNull(dtImprRows["IssueMethod"]) == "EXPIRY DATE")
            {
                strOrderBy = " Order By OrderField,[ExpiryDate] ";
                strOrderFile = ",case isnull([ExpiryDate],'1899-12-01') when '1899-12-01' then 1 else 0 end AS OrderField";
            }
            else if (CheckNull(dtImprRows["IssueMethod"]) == "MANUFACTURING DATE")
            {
                strOrderBy = " Order By OrderField,[ManufactureDate] ";
                strOrderFile = ",case isnull([ManufactureDate],'1899-12-01') when '1899-12-01' then 1 else 0 end AS OrderField";
            }
            string strDimColumnName = "";
            string strSql = "";
            if (CheckNull(dtImprRows["DimensionFlag"]) == "1")
            {
                strSql = "Select *" + strOrderFile + " from Impm1 Where StatusCode <> 'DEL' AND [DimensionFlag] = '1' AND [BalancePackingQty] > 0 AND [CustomerCode] = " + ReplaceWithNull(dtImgi1.Rows[0]["CustomerCode"]) + " " + strOrderBy;
                strDimColumnName = "BalancePackingQty";
            }
            else if (CheckNull(dtImprRows["DimensionFlag"]) == "2")
            {
                strSql = "Select *" + strOrderFile + " from Impm1 Where StatusCode <> 'DEL' AND [DimensionFlag] = '2' AND [BalanceWholeQty] > 0 AND [CustomerCode] = " + ReplaceWithNull(dtImgi1.Rows[0]["CustomerCode"]) + " " + strOrderBy;
                strDimColumnName = "BalanceWholeQty";
            }
            else
            {
                strSql = "Select *" + strOrderFile + " from Impm1 Where StatusCode <> 'DEL' AND [DimensionFlag] = '3' AND [BalanceLooseQty] > 0 AND [CustomerCode] = " + ReplaceWithNull(dtImgi1.Rows[0]["CustomerCode"]) + " " + strOrderBy;
                strDimColumnName = "BalanceLooseQty";
            }
            list.Add(strSql);
            list.Add(strDimColumnName);
            return list.ToArray();
        }

        private DataTable setSalesShipmentExportImgi2(DataTable dtImgi2, string[] strValueNew, DataTable dtImgi1)
        {
            string strProductCode = "-" + strValueNew[14] + "-" + strValueNew[18];
            DataTable dtImpr = GetSQLCommandReturnDTNew("Select * from Impr1 Where ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(dtImgi1.Rows[0]["CustomerCode"]));
            if (dtImpr == null || dtImpr.Rows.Count == 0) { return dtImgi2; }
            int intQty = CheckNullInt(strValueNew[15], 1);
            if (intQty <= 0) { return dtImgi2; }
            DataRow dtImprRows = dtImpr.Rows[0];
            string strSql = "";
            string[] strlist = getImpmSqlByDim(dtImprRows, dtImgi1);
            strSql = strlist[0];
            string strDimColumnName = strlist[1];
            while (intQty > 0)
            {
                DataTable dtImpm = GetSQLCommandReturnDTNew(strSql);
                if (dtImpm == null || dtImpm.Rows.Count == 0) { return dtImgi2; }
                int intDimQty = CheckNullInt(dtImpm.Rows[0][strDimColumnName], 1);
                if (intDimQty > intQty) { intDimQty = intQty; }
                DataRow dtImpmRows = dtImpm.Rows[0];
                dtImgi2 = PullProductByBatchTrxNo(CheckNullInt(dtImpm.Rows[0]["TrxNo"], 1), dtImgi2, dtImpmRows);
                DataRow dtImgi2Row = dtImgi2.Rows[dtImgi2.Rows.Count - 1];
                dtImgi2Row["SoNo"] = strValueNew[0];
                dtImgi2Row["SoLineItemNo"] = strValueNew[9];
                dtImgi2Row["StoreNo"] = strValueNew[1];
                dtImgi2Row["ProductDescription"] = strValueNew[19];
                dtImgi2Row["ProductCode"] = strProductCode;
                dtImgi2Row = SetFromImpmForImgi(dtImgi2Row, dtImgi2Row, dtImprRows, dtImpr, intDimQty);
                intQty = intQty - intDimQty;
            }


            return dtImgi2;
        }

        private DataRow SetFromImpmForImgi(DataRow dtImgi2Row, DataRow dtImpmRows, DataRow dtImprRows, DataTable dtImpr, int intDimQty)
        {
            if (dtImpr != null && dtImpr.Rows.Count > 0)
            {
                dtImgi2Row["CustomerCode"] = dtImprRows["CustomerCode"];
                dtImgi2Row["ProductTrxNo"] = dtImprRows["TrxNo"];
                dtImgi2Row["DimensionFlag"] = dtImprRows["DimensionFlag"];
                if (CheckNullInt(dtImprRows["DimensionFlag"], 1) == 1)
                {
                    dtImgi2Row["UomCode"] = dtImprRows["PackingUomCode"];
                    dtImgi2Row["PackingQty"] = intDimQty;
                    dtImgi2Row["Length"] = dtImprRows["PackingLength"];
                    dtImgi2Row["Width"] = dtImprRows["PackingWidth"];
                    dtImgi2Row["Height"] = dtImprRows["PackingHeight"];
                    dtImgi2Row["WholeQty"] = CheckNullInt(dtImgi2Row["PackingQty"], 1) * CheckNullInt(dtImprRows["PackingPackageSize"], 1);
                    dtImgi2Row["LooseQty"] = CheckNullInt(dtImgi2Row["WholeQty"], 1) * CheckNullInt(dtImprRows["WholePackageSize"], 1);
                    setImgi2Volumn("BalancePackingQty", intDimQty, dtImpmRows, ref dtImgi2Row);
                }
                else if (CheckNullInt(dtImprRows["DimensionFlag"], 1) == 2)
                {
                    dtImgi2Row["UomCode"] = dtImprRows["WholeUomCode"];
                    dtImgi2Row["WholeQty"] = intDimQty;
                    dtImgi2Row["Length"] = dtImprRows["WholeLength"];
                    dtImgi2Row["Width"] = dtImprRows["WholeWidth"];
                    dtImgi2Row["Height"] = dtImprRows["WholeHeight"];
                    dtImgi2Row["UnitVol"] = dtImprRows["WholeVolume"];
                    dtImgi2Row["UnitWt"] = dtImprRows["WholeWeight"];
                    //dtImgi2Row["QtyPerPallet"] = dtImprRows["WholePackageSize"];
                    dtImgi2Row["LooseQty"] = CheckNullInt(dtImgi2Row["WholeQty"], 1) * CheckNullInt(dtImprRows["WholePackageSize"], 1);
                    setImgi2Volumn("BalanceWholeQty", intDimQty, dtImpmRows, ref dtImgi2Row);
                }
                else
                {
                    dtImgi2Row["DimensionFlag"] = 3;
                    dtImgi2Row["UomCode"] = dtImprRows["LooseUomCode"];
                    dtImgi2Row["LooseQty"] = intDimQty;
                    dtImgi2Row["Length"] = dtImprRows["LooseLength"];
                    dtImgi2Row["Width"] = dtImprRows["LooseWidth"];
                    dtImgi2Row["Height"] = dtImprRows["LooseHeight"];
                    dtImgi2Row["UnitVol"] = dtImprRows["LooseVolume"];
                    dtImgi2Row["UnitWt"] = dtImprRows["LooseWeight"];
                    //dtImgi2Row["QtyPerPallet"] = CheckNullDouble(dtImprRows["PackingPackageSize"], 1) * CheckNullDouble(dtImprRows["WholePackageSize"], 1);
                    setImgi2Volumn("BalanceLooseQty", intDimQty, dtImpmRows, ref dtImgi2Row);
                }
                dtImgi2Row["Volume"] = CheckNullDouble(dtImgi2Row["Length"], 1) * CheckNullDouble(dtImgi2Row["Width"], 1) * CheckNullDouble(dtImgi2Row["Height"], 1);
                dtImgi2Row["ProductTrxNo"] = dtImprRows["TrxNo"];
            }
            return dtImgi2Row;
        }

        #endregion

        #region Sales Shipment Import

        private void setSalesShipmentImport()
        {
            string strFilter = "";
            if (this.m_strFilter1 != "") { strFilter = " AND " + m_strFilter1; }
            DataTable dtPur = GetSQLCommandReturnDTNew("Select Distinct SalesOrderNo from imgi1 Where imgi1.StatusCode = 'EXE' and (imgi1.EdiCount = 0 or imgi1.EdiCount is null) AND (Len(imgi1.SalesOrderNo)>0 OR imgi1.SalesOrderNo <> '') "+ strFilter);
            if (dtPur == null || dtPur.Rows.Count == 0) { return; }
            for (int i = 0; i < dtPur.Rows.Count; i++)
            {
                string strSql = "select Imgi1.SalesOrderNo,Imgi2.SoLineItemNo,Imgi2.StoreNo,Imgi2.ProductCode ,";
                strSql = strSql + "case Imgi2.DimensionFlag when '1' then Imgi2.PackingQty when '2' then Imgi2.WholeQty else Imgi2.LooseQty end AS DimensionQty,";
                strSql = strSql + "case Imgi2.DimensionFlag when '1' then(Select Impr1.PackingUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " when '2' then(Select Impr1.WholeUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " else  (Select Impr1.LooseUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo ) end AS UnitOfMeasureCode, ";
                strSql = strSql + " (Select Imso2.Volume From Imso2 where Imso2.TrxNo = (Select top 1  TrxNo from Imso1 Where Imso1.SalesOrderNo = Imgi1.SalesOrderNo ) AND SoLineItemNo = Imgi2.SoLineItemNo) AS QtyPerUOM,";
                strSql = strSql + "(Select Imso2.Qty From Imso2 where Imso2.TrxNo = (Select top 1  TrxNo from Imso1 Where Imso1.SalesOrderNo = Imgi1.SalesOrderNo ) AND SoLineItemNo = Imgi2.SoLineItemNo ) AS QtyToReceive,Imgi1.GoodsIssueNoteNo ";
                strSql = strSql + "  from Imgi1 Join Imgi2 on Imgi1.TrxNo = Imgi2.TrxNo Where Imgi1.StatusCode = 'EXE' and Imgi1.SalesOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]);
                DataTable dtRec = GetSQLCommandReturnDTNew(strSql);
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow < dtRec.Rows.Count; intRow++)
                    {
                        string[] strProduct = CheckNull(dtRec.Rows[intRow]["ProductCode"]).Split('-');
                        if (strProduct.Length == 3) { dtRec.Rows[intRow]["ProductCode"] = strProduct[1]; }
                        saveSaed2("", CheckNull(dtRec.Rows[intRow]["GoodsIssueNoteNo"]), CheckNull(dtRec.Rows[intRow]["SalesOrderNo"]), "SalesOrderNo", "ProductCode : " + dtRec.Rows[intRow]["ProductCode"] + " <" + dtRec.Rows[intRow]["DimensionQty"] + ">", "Success");
                    }
                    if (dtRec.Columns.Contains("GoodsIssueNoteNo")) { dtRec.Columns.Remove(dtRec.Columns["GoodsIssueNoteNo"]); }
                    string strFileName = this.m_strOUTFolder + @"\ReturnShip-" + CheckNull(dtPur.Rows[i][0]) + "-" + this.m_strYear + this.m_strMth + this.m_strDay + "-" + this.m_strHour + this.m_strMin + @".txt";
                    if (!Directory.Exists(@m_strOUTFolder)) { Directory.CreateDirectory(@m_strOUTFolder); }
                    this.saveToField(dtRec, @m_strOUTFolder, "|", "\"", false);
                    GetSQLCommandReturnIntNew("Update Imgi1 Set EdiCount = isnull(EdiCount,0)+1 Where Imgi1.StatusCode = 'EXE' and Imgi1.SalesOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]));
                }
            }
            saveLog("");
        }
        #endregion

        #region Sales Return Receipt Export
        private void UploadSalesReturnReceiptExport()
        {
            string[] strFileList = this.getFileList(".txt");
            string strErrorProductAll = "";
            string strErrorMessageAll = "";
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImgr1 = null;
                DataTable dtImgr2 = null;
                DataTable dtImpo1 = null;
                DataTable dtImpo2 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    if (dtImgr1 == null)
                    {
                        dtImgr1 = GetSQLCommandReturnDTNew("Select Top 0 * from imgr1");
                        dtImgr2 = GetSQLCommandReturnDTNew("Select Top 0 * from imgr2");
                        dtImpo1 = GetSQLCommandReturnDTNew("Select Top 0 * from impo1");
                        dtImpo2 = GetSQLCommandReturnDTNew("Select Top 0 * from impo2");
                    }
                    string strErrorProduct = "";
                    string strCurrentProduct = "";
                    string strErrorMessage = "";
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        List<string> strPurchaseOrderNoList = new List<string> { };
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 35)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 35") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Purchase Order No", "txt file format error-> Column Count <> 35", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 35" + "|";
                                    continue;
                                }
                            }
                            try
                            {
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                if (this.getCheckErrorProductCode("-" + strValue[14] + "-" + strValue[18]))
                                {
                                    if (strErrorProduct.IndexOf("-" + strValue[14] + "-" + strValue[18]) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strValue[0], "Purchase Order No", strValue[0], "Fail");
                                        strErrorProduct = strErrorProduct + "-" + strValue[14] + "-" + strValue[18] + "|";
                                    }
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0)
                                {
                                    saveSaed2(strFileList[intI], "", strValue[0], "Purchase Order No", ex.Message, "Fail");
                                    strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|";
                                }
                            }
                        }
                        if (strErrorProduct != "")
                        {
                            strErrorProduct = strErrorProduct.Substring(0, strErrorProduct.Length - 1);
                            strErrorProductAll = strErrorProductAll + "\r\n" + strFileList[intI] + "\r\n Below Product not in sysfreight Product master (Customer : " + this.strDefaultCustomerCode + "), and not upload.\r\n" + strErrorProduct + "\r\n";
                            continue;
                        }
                        else if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + strErrorMessage + "\r\n";
                            continue;
                        }
                        strPurchaseOrderNoList.Clear();
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            try
                            {
                                string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                strCurrentProduct = "-" + strValue[12] + "-" + strValue[16];
                                strPurchaseOrderNoList.Add(strValue[0]);
                                if (CheckUploadData("Imgr1", "PurchaseOrderNo", strValue[0]))
                                {
                                    if (strErrorMessage.IndexOf("Purchase Order No:" + strValue[0] + " already exist") < 0) { strErrorMessage = strErrorMessage + "Purchase Order No:" + strValue[0] + " already exist | "; }
                                    continue;
                                }
                                if (strValue != null && strValue.Length == 35)
                                {
                                    dtImgr1.Rows.Clear();
                                    dtImgr2.Rows.Clear();
                                    dtImpo1.Rows.Clear();
                                    dtImpo2.Rows.Clear();
                                    dtImgr1.Rows.Add(dtImgr1.NewRow());
                                    dtImpo1.Rows.Add(dtImpo1.NewRow());
                                    dtImgr1.Rows[0]["PurchaseOrderNo"] = strValue[0];
                                    dtImpo1.Rows[0]["PurchaseOrderNo"] = strValue[0];
                                    dtImgr1.Rows[0]["ReceiptDate"] = this.CheckNullDate(strValue[3], 2);
                                    dtImpo1.Rows[0]["OrderDate"] = dtImgr1.Rows[0]["ReceiptDate"];
                                    dtImgr1.Rows[0]["WarehouseCode"] = this.strDefaultWarehouse;
                                    dtImgr1 = setDefaultCustomerNameAddress(dtImgr1);
                                    dtImpo1 = this.setDefaultVenderByCustomer(dtImpo1);
                                    int intImgrTrxNo = -1;
                                    int intImpoTrxNo = -1;
                                    for (int intDetail = intJ; intDetail < strValueList.Length; intDetail++)
                                    {
                                        string[] strValueNew = this.getLineDetail(strValueList[intDetail], '|');
                                        if (strValueNew != null && strValueNew.Length >= 35)
                                        {
                                            if (strValue[0] == strValueNew[0])
                                            {
                                                dtImgr2.Rows.Add(dtImgr2.NewRow());
                                                dtImpo2.Rows.Add(dtImpo2.NewRow());
                                                dtImgr2 = setSalesReturnReceiptExportImgr2(dtImgr2, strValueNew, dtImgr1);
                                                dtImpo2 = setSalesReturnReceiptExportImpo2(dtImpo2, strValueNew);
                                            }
                                        }
                                    }
                                    dtImgr1 = UpdateImgrTotal(dtImgr1, dtImgr2);
                                    intImgrTrxNo = this.getTrxNoSaveGenerateNumber("GoodsReceiptNoteNo", "imgr1", "NextGoodsReceiptNo", dtImgr1);
                                    intImpoTrxNo = this.GetTrxNoSaveDateRemoveTrxNo("Impo1", dtImpo1);
                                    if (intImgrTrxNo > 0 && dtImgr2 != null && dtImgr2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dtImgr2.Rows.Count; i++)
                                        {
                                            dtImgr2.Rows[i]["TrxNo"] = intImgrTrxNo;
                                            dtImgr2.Rows[i]["LineItemNo"] = i + 1;
                                        }
                                        this.InsertTableRecordByDatatableNew("Imgr2", dtImgr2, false);
                                        for (int i = 0; i < dtImgr2.Rows.Count; i++)
                                        {
                                            setUpdateImpm1AndImpr1ByImgr(dtImgr2.Rows[i]);
                                            saveSaed2(strFileList[intI], CheckNull(dtImgr1.Rows[0]["GoodsReceiptNoteNo"]), CheckNull(dtImgr1.Rows[0]["PurchaseOrderNo"]), "PurchaseOrderNo", "ProductCode : " + dtImgr2.Rows[i]["ProductCode"] + " <" + getDimensionQty(dtImgr2.Rows[i]) + ">", "Success");
                                        }
                                    }
                                    if (intImpoTrxNo > 0 && dtImpo2 != null && dtImpo2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dtImpo2.Rows.Count; i++)
                                        {
                                            dtImpo2.Rows[i]["TrxNo"] = intImpoTrxNo;
                                            dtImpo2.Rows[i]["LineItemNo"] = i + 1;
                                        }
                                        this.InsertTableRecordByDatatableNew("Impo2", dtImpo2, false);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0) { strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|"; }
                            }
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorProductAll + strErrorMessageAll);
        }

        private DataTable setSalesReturnReceiptExportImpo2(DataTable dtImpo2, string[] strValueNew)
        {
            string strProductCode = "-" + strValueNew[14] + "-" + strValueNew[18];
            var dtImpo2Row = dtImpo2.Rows[dtImpo2.Rows.Count - 1];
            dtImpo2Row["UomCode"] = strValueNew[16];
            dtImpo2Row["Qty"] = CheckNullInt(strValueNew[15], 1);
            dtImpo2Row["Description"] = strValueNew[19];
            dtImpo2Row["ProductCode"] = strProductCode;
            dtImpo2Row["PoLineItemNo"] = strValueNew[9];
            dtImpo2Row["Volume"] = CheckNullDouble(strValueNew[17], 1);
            return dtImpo2;
        }


        private DataRow getImgr2FromImpr(DataRow dtImgr2Row, DataTable dtImpr, int intDimQty)
        {
            if (dtImpr != null && dtImpr.Rows.Count > 0)
            {
                DataRow dtImprRow = dtImpr.Rows[0];
                dtImgr2Row["CustomerCode"] = dtImprRow["CustomerCode"];
                dtImgr2Row["ProductTrxNo"] = dtImprRow["TrxNo"];
                dtImgr2Row["DimensionFlag"] = dtImprRow["DimensionFlag"];
                if (CheckNullInt(dtImprRow["DimensionFlag"], 1) == 1)
                {
                    dtImgr2Row["UomCode"] = dtImprRow["PackingUomCode"];
                    dtImgr2Row["PackingQty"] = intDimQty;
                    dtImgr2Row["Length"] = dtImprRow["PackingLength"];
                    dtImgr2Row["Width"] = dtImprRow["PackingWidth"];
                    dtImgr2Row["Height"] = dtImprRow["PackingHeight"];
                    dtImgr2Row["UnitVol"] = dtImprRow["PackingVolume"];
                    dtImgr2Row["UnitWt"] = dtImprRow["PackingWeight"];
                    dtImgr2Row["QtyPerPallet"] = 1;
                    dtImgr2Row["WholeQty"] = CheckNullInt(dtImgr2Row["PackingQty"], 1) * CheckNullInt(dtImprRow["PackingPackageSize"], 1);
                    dtImgr2Row["LooseQty"] = CheckNullInt(dtImgr2Row["WholeQty"], 1) * CheckNullInt(dtImprRow["WholePackageSize"], 1);
                }
                else if (CheckNullInt(dtImpr.Rows[0]["DimensionFlag"], 1) == 2)
                {
                    dtImgr2Row["UomCode"] = dtImpr.Rows[0]["WholeUomCode"];
                    dtImgr2Row["WholeQty"] = intDimQty;
                    dtImgr2Row["LooseQty"] = CheckNullInt(dtImgr2Row["WholeQty"], 1) * CheckNullInt(dtImprRow["WholePackageSize"], 1);
                    dtImgr2Row["Length"] = dtImpr.Rows[0]["WholeLength"];
                    dtImgr2Row["Width"] = dtImpr.Rows[0]["WholeWidth"];
                    dtImgr2Row["Height"] = dtImpr.Rows[0]["WholeHeight"];
                    dtImgr2Row["UnitVol"] = dtImpr.Rows[0]["WholeVolume"];
                    dtImgr2Row["UnitWt"] = dtImpr.Rows[0]["WholeWeight"];
                    dtImgr2Row["QtyPerPallet"] = dtImpr.Rows[0]["WholePackageSize"];
                }
                else
                {
                    dtImgr2Row["DimensionFlag"] = 3;
                    dtImgr2Row["UomCode"] = dtImpr.Rows[0]["LooseUomCode"];
                    dtImgr2Row["LooseQty"] = intDimQty;
                    dtImgr2Row["Length"] = dtImpr.Rows[0]["LooseLength"];
                    dtImgr2Row["Width"] = dtImpr.Rows[0]["LooseWidth"];
                    dtImgr2Row["Height"] = dtImpr.Rows[0]["LooseHeight"];
                    dtImgr2Row["UnitVol"] = dtImpr.Rows[0]["LooseVolume"];
                    dtImgr2Row["UnitWt"] = dtImpr.Rows[0]["LooseWeight"];
                    dtImgr2Row["QtyPerPallet"] = CheckNullDouble(dtImpr.Rows[0]["PackingPackageSize"], 1) * CheckNullDouble(dtImpr.Rows[0]["WholePackageSize"], 1);
                }
                dtImgr2Row["Volume"] = CheckNullDouble(dtImgr2Row["Length"], 1) * CheckNullDouble(dtImgr2Row["Width"], 1) * CheckNullDouble(dtImgr2Row["Height"], 1);
                dtImgr2Row["ProductTrxNo"] = dtImpr.Rows[0]["TrxNo"];
            }
            return dtImgr2Row;
        }
        private DataTable setSalesReturnReceiptExportImgr2(DataTable dtImgr2, string[] strValueNew, DataTable dtImgr1)
        {
            DataRow dtImgr2Row = dtImgr2.Rows[dtImgr2.Rows.Count - 1];
            string strProductCode = "-" + strValueNew[14] + "-" + strValueNew[18];
            dtImgr2Row["WarehouseCode"] = strDefaultWarehouse;
            dtImgr2Row["StoreNo"] = strDefaultStoreNo;
            dtImgr2Row["PoNo"] = strValueNew[0];
            dtImgr2Row["UserDefine1"] = strValueNew[0];
            dtImgr2Row["PoLineItemNo"] = strValueNew[9];
            dtImgr2Row["StoreLocation"] = strValueNew[13];
            dtImgr2Row["ProductDescription"] = strValueNew[19];
            dtImgr2Row["ProductCode"] = strProductCode;
            DataTable dtImpr = GetSQLCommandReturnDTNew("Select * from Impr1 Where ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(dtImgr1.Rows[0]["CustomerCode"]));
            dtImgr2Row = this.getImgr2FromImpr(dtImgr2Row, dtImpr, CheckNullInt(strValueNew[15], 1));
            return dtImgr2;
        }

        #endregion

        #region Sales Return Receipt Import
        private void setSalesReturnReceiptImport()
        {
            string strFilter = "";
            if (this.m_strFilter1 != "") { strFilter = " AND " + m_strFilter1; }
            DataTable dtPur = GetSQLCommandReturnDTNew("Select Distinct PurchaseOrderNo from imgr1 Where Imgr1.StatusCode = 'EXE' and (Imgr1.EdiCount = 0 or Imgr1.EdiCount is null) AND (Len(Imgr1.PurchaseOrderNo)>0 OR Imgr1.PurchaseOrderNo <> '') "+ strFilter);
            if (dtPur == null || dtPur.Rows.Count == 0) { return; }
            for (int i = 0; i < dtPur.Rows.Count; i++)
            {
                string strSql = "select Imgr1.PurchaseOrderNo,Imgr2.PoLineItemNo,Imgr2.StoreLocation,Imgr2.ProductCode ,";
                strSql = strSql + "case Imgr2.DimensionFlag when '1' then imgr2.PackingQty when '2' then imgr2.WholeQty else imgr2.LooseQty end AS DimensionQty,";
                strSql = strSql + "case Imgr2.DimensionFlag when '1' then(Select Impr1.PackingUomCode from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "when '2' then(Select Impr1.WholeUomCode from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "else  (Select Impr1.LooseUomCode from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) end AS UnitOfMeasureCode,  ";
                strSql = strSql + "case Imgr2.DimensionFlag when '1' then(Select Impr1.PackingPackageSize * Impr1.WholePackageSize from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "when '2' then(Select Impr1.WholePackageSize from impr1 where impr1.TrxNo = imgr2.ProductTrxNo ) ";
                strSql = strSql + "else  1 end AS QtyPerUnitOfMeasure, ";
                strSql = strSql + "(Select Impo2.Volume From impo2 where Impo2.TrxNo = (Select top 1  TrxNo from impo1 Where Impo1.PurchaseOrderNo = imgr1.PurchaseOrderNo ) AND PoLineItemNo = Imgr2.PoLineItemNo) AS QtyPerUOM, ";
                strSql = strSql + "(Select Impo2.Qty From impo2 where Impo2.TrxNo = (Select top 1  TrxNo from impo1 Where Impo1.PurchaseOrderNo = imgr1.PurchaseOrderNo ) AND PoLineItemNo = Imgr2.PoLineItemNo ) AS QtyToReceive,Imgr1.ReceiptDate,Imgr1.GoodsReceiptNoteNo ";
                strSql = strSql + " from imgr1 Join imgr2 on Imgr1.TrxNo = imgr2.TrxNo Where Imgr1.StatusCode = 'EXE' and Imgr1.PurchaseOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]);
                DataTable dtRec = GetSQLCommandReturnDTNew(strSql);
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow < dtRec.Rows.Count; intRow++)
                    {
                        string[] strProduct = CheckNull(dtRec.Rows[intRow]["ProductCode"]).Split('-');
                        if (strProduct.Length == 3) { dtRec.Rows[intRow]["ProductCode"] = strProduct[1]; }
                        saveSaed2("", CheckNull(dtRec.Rows[intRow]["GoodsReceiptNoteNo"]), CheckNull(dtRec.Rows[intRow]["PurchaseOrderNo"]), "PurchaseOrderNo", "ProductCode : " + dtRec.Rows[intRow]["ProductCode"] + " <" + dtRec.Rows[intRow]["DimensionQty"] + ">", "Success");
                    }
                    if (dtRec.Columns.Contains("GoodsReceiptNoteNo")) { dtRec.Columns.Remove(dtRec.Columns["GoodsReceiptNoteNo"]); }
                    string strFileName = this.m_strOUTFolder + @"\PurchRcpt-" + CheckNull(dtPur.Rows[i][0]) + "-" + this.m_strYear + this.m_strMth + this.m_strDay + "-" + this.m_strHour + this.m_strMin + @".txt";
                    if (!Directory.Exists(@m_strOUTFolder)) { Directory.CreateDirectory(@m_strOUTFolder); }
                    this.saveToField(dtRec, @m_strOUTFolder, "|", "\"", false);
                    GetSQLCommandReturnIntNew("Update Imgr1 Set EdiCount = isnull(EdiCount,0)+1 Where Imgr1.StatusCode = 'EXE' and Imgr1.PurchaseOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]));
                }
            }
            saveLog("");
        }
        #endregion

        #region Transfer Shipment Export
        private void UploadTransferShipmentExport()
        {
            string[] strFileList = this.getFileList(".txt");
            string strErrorProductAll = "";
            string strErrorMessageAll = "";
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImgi1 = null;
                DataTable dtImgi2 = null;
                DataTable dtImso1 = null;
                DataTable dtImso2 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    if (dtImgi1 == null)
                    {
                        dtImgi1 = GetSQLCommandReturnDTNew("Select Top 0 * from imgi1");
                        dtImgi2 = GetSQLCommandReturnDTNew("Select Top 0 * from imgi2");
                        dtImso1 = GetSQLCommandReturnDTNew("Select Top 0 * from imso1");
                        dtImso2 = GetSQLCommandReturnDTNew("Select Top 0 * from imso2");
                    }
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        string strErrorProduct = "";
                        string strCurrentProduct = "";
                        string strErrorMessage = "";
                        List<string> strPurchaseOrderNoList = new List<string> { };
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 36)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 36") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Sales Order No", "txt file format error-> Column Count <> 36", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 36" + "|";
                                    continue;
                                }
                            }
                            try
                            {
                                if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                                if (this.getCheckErrorProductCode("-" + strValue[13] + "-" + strValue[17]))
                                {
                                    if (strErrorProduct.IndexOf("-" + strValue[13] + "-" + strValue[17]) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", strValue[0], "Fail");
                                        strErrorProduct = strErrorProduct + "-" + strValue[13] + "-" + strValue[17] + "|";
                                    }
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (strErrorMessage.IndexOf(strCurrentProduct + ":" + ex.Message) < 0)
                                {
                                    saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", ex.Message, "Fail");
                                    strErrorMessage = strErrorMessage + strCurrentProduct + ":" + ex.Message + "|";
                                }
                            }
                        }
                        if (strErrorProduct != "")
                        {
                            strErrorProduct = strErrorProduct.Substring(0, strErrorProduct.Length - 1);
                            strErrorProductAll = strErrorProductAll + "\r\n" + strFileList[intI] + "\r\n Below Product not in sysfreight Product master (Customer : " + this.strDefaultCustomerCode + "), and not upload.\r\n" + strErrorProduct + "\r\n";
                            continue;
                        }
                        else if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + strErrorMessage + "\r\n";
                            continue;
                        }
                        strPurchaseOrderNoList.Clear();
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strPurchaseOrderNoList.Contains(strValue[0])) { continue; }
                            strPurchaseOrderNoList.Add(strValue[0]);
                            if (CheckUploadData("Imgi1", "SalesOrderNo", strValue[0]))
                            {
                                if (strErrorMessage.IndexOf("Sales Order No:" + strValue[0] + " already exist") < 0) { strErrorMessage = strErrorMessage + "Sales Order No:" + strValue[0] + " already exist | "; }
                                continue;
                            }
                            if (strValue != null && strValue.Length == 36)
                            {
                                dtImgi1.Rows.Clear();
                                dtImgi2.Rows.Clear();
                                dtImso1.Rows.Clear();
                                dtImso2.Rows.Clear();
                                dtImgi1.Rows.Add(dtImgi1.NewRow());
                                dtImso1.Rows.Add(dtImso1.NewRow());
                                dtImgi1.Rows[0]["SalesOrderNo"] = strValue[0];
                                dtImso1.Rows[0]["SalesOrderNo"] = strValue[0];
                                dtImgi1.Rows[0]["IssueDateTime"] = this.CheckNullDate(strValue[2], 2);
                                dtImso1.Rows[0]["OrderDate"] = dtImgi1.Rows[0]["IssueDateTime"];
                                dtImso1.Rows[0]["IssueTo"] = strValue[7];
                                dtImso1.Rows[0]["CollectFromAddress1"] = strValue[25];
                                dtImso1.Rows[0]["CollectFromAddress2"] = strValue[26];
                                dtImso1.Rows[0]["DeliveryToAddress1"] = strValue[28];
                                dtImso1.Rows[0]["DeliveryToAddress2"] = strValue[29];
                                dtImso1.Rows[0]["Description1"] = strValue[31].Substring(0, 50);
                                if (strValue[31].Length > 50) { dtImso1.Rows[0]["Description2"] = strValue[31].Substring(50, strValue[31].Length - 50); }
                                dtImso1.Rows[0]["DeliveryToContactName"] = strValue[35];
                                dtImso1.Rows[0]["DeliveryToAddress3"] = strValue[32];
                                dtImso1.Rows[0]["DeliveryToAddress4"] = strValue[33];
                                dtImgi1 = setDefaultCustomerNameAddress(dtImgi1);
                                dtImso1 = this.setDefaultVenderByCustomer(dtImso1);
                                int intImgiTrxNo = -1;
                                int intImsoTrxNo = -1;
                                for (int intDetail = intJ; intDetail < strValueList.Length; intDetail++)
                                {
                                    string[] strValueNew = this.getLineDetail(strValueList[intDetail], '|');
                                    if (strValueNew != null && strValueNew.Length == 36)
                                    {
                                        if (strValue[0] == strValueNew[0])
                                        {
                                            //dtImgi2.Rows.Add(dtImgi2.NewRow());
                                            dtImso2.Rows.Add(dtImso2.NewRow());
                                            dtImgi2 = setTransferShipmentExportImgi2(dtImgi2, strValueNew, dtImgi1);
                                            dtImso2 = setTransferShipmentExportImso2(dtImso2, strValueNew);
                                        }
                                    }
                                }
                                if (dtImgi2 == null || dtImgi2.Rows.Count == 0)
                                {
                                    strErrorMessageAll = strErrorMessageAll + "\r\n" + strFileList[intI] + "\r\n" + "Sales Order No : " + strValue[0] + ", not Dim Qty to Issue." + "\r\n";
                                    saveSaed2(strFileList[intI], "", strValue[0], "Sales Order No", strValue[0], "Fail");
                                    continue;
                                }
                                dtImgi1 = UpdateImgiTotal(dtImgi1, dtImgi2);
                                intImgiTrxNo = this.getTrxNoSaveGenerateNumber("GoodsIssueNoteNo", "imgi1", "NextGoodsIssueNo", dtImgi1);
                                intImsoTrxNo = this.GetTrxNoSaveDateRemoveTrxNo("Imso1", dtImso1);
                                if (intImgiTrxNo > 0 && dtImgi2 != null && dtImgi2.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtImgi2.Rows.Count; i++)
                                    {
                                        dtImgi2.Rows[i]["TrxNo"] = intImgiTrxNo;
                                        dtImgi2.Rows[i]["LineItemNo"] = i + 1;
                                    }
                                    this.InsertTableRecordByDatatableNew("Imgi2", dtImgi2, false);
                                    for (int i = 0; i < dtImgi2.Rows.Count; i++)
                                    {
                                        setUpdateImpm1AndImpr1ByImgi(dtImgi2.Rows[i]);
                                        saveSaed2(strFileList[intI], CheckNull(dtImgi1.Rows[0]["GoodsIssueNoteNo"]), CheckNull(dtImgi1.Rows[0]["SalesOrderNo"]), "SalesOrderNo", "ProductCode : " + dtImgi2.Rows[i]["ProductCode"] + " <" + getDimensionQty(dtImgi2.Rows[i]) + ">", "Success");
                                    }

                                }
                                if (intImsoTrxNo > 0 && dtImso2 != null && dtImso2.Rows.Count > 0)
                                {
                                    for (int i = 0; i < dtImso2.Rows.Count; i++)
                                    {
                                        dtImso2.Rows[i]["TrxNo"] = intImsoTrxNo;
                                        dtImso2.Rows[i]["LineItemNo"] = i + 1;
                                    }
                                    this.InsertTableRecordByDatatableNew("Imso2", dtImso2, false);
                                }
                            }
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorProductAll + strErrorMessageAll);
        }

        private DataTable setTransferShipmentExportImso2(DataTable dtImso2, string[] strValueNew)
        {
            string strProductCode = "-" + strValueNew[13] + "-" + strValueNew[17];
            var dtImso2Row = dtImso2.Rows[dtImso2.Rows.Count - 1];
            dtImso2Row["UomCode"] = strValueNew[15];
            dtImso2Row["Description"] = strValueNew[18];
            dtImso2Row["ProductCode"] = strProductCode;
            dtImso2Row["SoLineItemNo"] = strValueNew[8];
            dtImso2Row["Volume"] = CheckNullDouble(strValueNew[16], 1);
            return dtImso2;
        }

        private DataTable setTransferShipmentExportImgi2(DataTable dtImgi2, string[] strValueNew, DataTable dtImgi1)
        {
            string strProductCode = "-" + strValueNew[13] + "-" + strValueNew[17];
            DataTable dtImpr = GetSQLCommandReturnDTNew("Select * from Impr1 Where ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(dtImgi1.Rows[0]["CustomerCode"]));
            if (dtImpr == null || dtImpr.Rows.Count == 0) { return dtImgi2; }
            int intQty = CheckNullInt(strValueNew[14], 1);
            if (intQty <= 0) { return dtImgi2; }
            DataRow dtImprRows = dtImpr.Rows[0];
            string strSql = "";
            string[] strlist = getImpmSqlByDim(dtImprRows, dtImgi1);
            strSql = strlist[0];
            string strDimColumnName = strlist[1];
            while (intQty > 0)
            {
                DataTable dtImpm = GetSQLCommandReturnDTNew(strSql);
                if (dtImpm == null || dtImpm.Rows.Count == 0) { return dtImgi2; }
                int intDimQty = CheckNullInt(dtImpm.Rows[0][strDimColumnName], 1);
                if (intDimQty > intQty) { intDimQty = intQty; }
                DataRow dtImpmRows = dtImpm.Rows[0];
                dtImgi2 = PullProductByBatchTrxNo(CheckNullInt(dtImpm.Rows[0]["TrxNo"], 1), dtImgi2, dtImpmRows);
                DataRow dtImgi2Row = dtImgi2.Rows[dtImgi2.Rows.Count - 1];
                dtImgi2Row["SoNo"] = strValueNew[0];
                dtImgi2Row["SoLineItemNo"] = strValueNew[8];
                dtImgi2Row["StoreNo"] = strValueNew[1];
                dtImgi2Row["ProductDescription"] = strValueNew[18];
                dtImgi2Row["ProductCode"] = strProductCode;
                dtImgi2Row = SetFromImpmForImgi(dtImgi2Row, dtImgi2Row, dtImprRows, dtImpr, intDimQty);
                intQty = intQty - intDimQty;
            }
            return dtImgi2;
        }

        #endregion

        #region Transfer Shipment Import
        private void setTransferShipmentImport()
        {
            string strFilter = "";
            if (this.m_strFilter1 != "") { strFilter = " AND " + m_strFilter1; }
            DataTable dtPur = GetSQLCommandReturnDTNew("Select Distinct SalesOrderNo from imgi1 Where imgi1.StatusCode = 'EXE' and (imgi1.EdiCount = 0 or imgi1.EdiCount is null) AND (Len(imgi1.SalesOrderNo)>0 OR imgi1.SalesOrderNo <> '') "+ strFilter);
            if (dtPur == null || dtPur.Rows.Count == 0) { return; }
            for (int i = 0; i < dtPur.Rows.Count; i++)
            {
                string strSql = "select Imgi1.SalesOrderNo,Imgi2.SoLineItemNo,Imgi2.StoreNo,Imgi2.ProductCode ,";
                strSql = strSql + "case Imgi2.DimensionFlag when '1' then Imgi2.PackingQty when '2' then Imgi2.WholeQty else Imgi2.LooseQty end AS DimensionQty,";
                strSql = strSql + "case Imgi2.DimensionFlag when '1' then(Select Impr1.PackingUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " when '2' then(Select Impr1.WholeUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo )";
                strSql = strSql + " else  (Select Impr1.LooseUomCode from impr1 where impr1.TrxNo = Imgi2.ProductTrxNo ) end AS UnitOfMeasureCode, ";
                strSql = strSql + " (Select Imso2.Volume From Imso2 where Imso2.TrxNo = (Select top 1  TrxNo from Imso1 Where Imso1.SalesOrderNo = Imgi1.SalesOrderNo ) AND SoLineItemNo = Imgi2.SoLineItemNo) AS QtyPerUOM,";
                strSql = strSql + "(Select Imso2.Qty From Imso2 where Imso2.TrxNo = (Select top 1  TrxNo from Imso1 Where Imso1.SalesOrderNo = Imgi1.SalesOrderNo ) AND SoLineItemNo = Imgi2.SoLineItemNo ) AS QtyToReceive,Imgi1.GoodsIssueNoteNo ";
                strSql = strSql + "  from Imgi1 Join Imgi2 on Imgi1.TrxNo = Imgi2.TrxNo Where Imgi1.StatusCode = 'EXE' and Imgi1.SalesOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]);
                DataTable dtRec = GetSQLCommandReturnDTNew(strSql);
                if (dtRec != null && dtRec.Rows.Count > 0)
                {
                    for (int intRow = 0; intRow < dtRec.Rows.Count; intRow++)
                    {
                        string[] strProduct = CheckNull(dtRec.Rows[intRow]["ProductCode"]).Split('-');
                        if (strProduct.Length == 3) { dtRec.Rows[intRow]["ProductCode"] = strProduct[1]; }
                        saveSaed2("", CheckNull(dtRec.Rows[intRow]["GoodsIssueNoteNo"]), CheckNull(dtRec.Rows[intRow]["SalesOrderNo"]), "SalesOrderNo", "ProductCode : " + dtRec.Rows[intRow]["ProductCode"] + " <" + dtRec.Rows[intRow]["DimensionQty"] + ">", "Success");
                    }
                    if (dtRec.Columns.Contains("GoodsIssueNoteNo")) { dtRec.Columns.Remove(dtRec.Columns["GoodsIssueNoteNo"]); }
                    string strFileName = this.m_strOUTFolder + @"\ReturnShip-" + CheckNull(dtPur.Rows[i][0]) + "-" + this.m_strYear + this.m_strMth + this.m_strDay + "-" + this.m_strHour + this.m_strMin + @".txt";
                    if (!Directory.Exists(@m_strOUTFolder)) { Directory.CreateDirectory(@m_strOUTFolder); }
                    this.saveToField(dtRec, @m_strOUTFolder, "|", "\"", false);
                    GetSQLCommandReturnIntNew("Update Imgi1 Set EdiCount = isnull(EdiCount,0)+1 Where Imgi1.StatusCode = 'EXE' and Imgi1.SalesOrderNo = " + ReplaceWithNull(dtPur.Rows[i][0]));
                }
            }
            saveLog("");
        }
        #endregion

        #region Item Variant

        private void UploadItemVariant()
        {
            string strErrorMessageAll = "";
            string[] strFileList = this.getFileList(".txt");
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImpr1 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    string strErrorMessage = "";
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 11)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 11") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Product Code", "txt file format error-> Column Count <> 11", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 11" + "|";
                                    continue;
                                }
                            } 
                            string strProductCode = strValue[10];
                            if (strProductCode.Split('-').Length == 3)
                            {
                                strProductCode = "-" + strProductCode.Split('|')[1] + "-" + strProductCode.Split('|')[2];
                            }
                            if (strValue != null && strValue.Length == 11)
                            {
                                try
                                {
                                    dtImpr1 = GetSQLCommandReturnDTNew("Select Top 1 * from impr1 ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(this.strDefaultCustomerCode));
                                    if (dtImpr1 != null && dtImpr1.Rows.Count > 0)
                                    {
                                        GetSQLCommandReturnIntNew("Update Impr1 Set ProductDescription = " + ReplaceWithNull(strValue[2]) + ",BrandName = " + ReplaceWithNull(strValue[4]) + ",Model = " + ReplaceWithNull(strValue[0]) + " Where TrxNo = " + ReplaceWithNull(dtImpr1.Rows[0]["TrxNo"], 1));
                                    }
                                    else
                                    {
                                        dtImpr1 = GetSQLCommandReturnDTNew("Select Top 0 * from impr1");
                                        dtImpr1.Rows.Add(dtImpr1.NewRow());
                                        DataRow dtImprRow = dtImpr1.Rows[0];
                                        dtImprRow["ProductCode"] = strProductCode;
                                        dtImprRow["ProductDescription"] = strValue[2];
                                        dtImprRow["BrandName"] = strValue[4];
                                        dtImprRow["CustomerCode"] = strDefaultCustomerCode;
                                        dtImprRow["Model"] = strValue[0];
                                        this.InsertTableRecordByDatatableNew("Impr1", dtImpr1, true);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (strErrorMessage.IndexOf(strProductCode + ":" + ex.Message) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strProductCode, "Product Code", ex.Message, "Fail");
                                        strErrorMessage = strErrorMessage + strProductCode + ":" + ex.Message + "|";
                                    }
                                }
                            }                           
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorMessageAll);
        }
        #endregion

        #region Master item 

        #endregion               

        #region Item Barcodes
        private void UploadItemBarcodes()
        {
            string strErrorMessageAll = "";
            string[] strFileList = this.getFileList(".txt");
            if (strFileList != null && strFileList.Length > 0)
            {
                DataTable dtImpr1 = null;
                for (int intI = 0; intI < strFileList.Length; intI++)
                {
                    string strErrorMessage = "";
                    string[] strValueList = this.getReadFromTXTFile(strFileList[intI]);
                    if (strValueList != null && strValueList.Length > 0)
                    {
                        for (int intJ = 0; intJ < strValueList.Length; intJ++)
                        {
                            string[] strValue = this.getLineDetail(strValueList[intJ], '|');
                            if (strValue == null || strValue.Length != 6)
                            {
                                if (strErrorMessage.IndexOf(strFileList[intI] + ":" + "txt file format error-> Column Count <> 11") < 0)
                                {
                                    saveSaed2(strFileList[intI], "", "", "Product Code", "txt file format error-> Column Count <> 11", "Fail");
                                    strErrorMessage = strErrorMessage + strFileList[intI] + ":" + "txt file format error-> Column Count <> 11" + "|";
                                    continue;
                                }
                            }
                            string strProductCode ="-" + strValue[0] + "-" + strValue[4];                            
                            if (strValue != null && strValue.Length == 6)
                            {
                                try
                                {
                                    dtImpr1 = GetSQLCommandReturnDTNew("Select Top 1 * from impr1 ProductCode = " + ReplaceWithNull(strProductCode) + " AND CustomerCode = " + ReplaceWithNull(this.strDefaultCustomerCode));
                                    if (dtImpr1 != null && dtImpr1.Rows.Count > 0)
                                    {
                                        if (CheckNull(dtImpr1.Rows[0]["UserDefine1"]) == "")
                                        {
                                            GetSQLCommandReturnIntNew("Update Impr1 Set UserDefine1=" + ReplaceWithNull(strValue[1]) + " Where TrxNo = " + ReplaceWithNull(dtImpr1.Rows[0]["TrxNo"], 1));
                                        }
                                        else
                                        {
                                            GetSQLCommandReturnIntNew("Update Impr1 Set UserDefine11=" + ReplaceWithNull(strValue[1]) + " Where TrxNo = " + ReplaceWithNull(dtImpr1.Rows[0]["TrxNo"], 1));
                                        }
                                    }
                                    else
                                    {
                                        dtImpr1 = GetSQLCommandReturnDTNew("Select Top 0 * from impr1");
                                        dtImpr1.Rows.Add(dtImpr1.NewRow());
                                        DataRow dtImprRow = dtImpr1.Rows[0];
                                        dtImprRow["ProductCode"] = strProductCode;
                                        dtImprRow["ProductDescription"] = strProductCode;
                                        dtImprRow["CustomerCode"] = strDefaultCustomerCode;
                                        dtImprRow["UserDefine1"] = strValue[1];
                                        this.InsertTableRecordByDatatableNew("Impr1", dtImpr1, true);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (strErrorMessage.IndexOf(strProductCode + ":" + ex.Message) < 0)
                                    {
                                        saveSaed2(strFileList[intI], "", strProductCode, "Product Code", ex.Message, "Fail");
                                        strErrorMessage = strErrorMessage + strProductCode + ":" + ex.Message + "|";
                                    }
                                }
                            }
                            else
                            {
                                saveSaed2(strFileList[intI], "", strProductCode, "Product Code", "txt file format error-> Column Count <> 11", "Fail");
                                strErrorMessage = strErrorMessage + strProductCode + ":" + "txt file format error-> Column Count <> 11" + "|";
                                return;
                            }
                        }
                        if (strErrorMessage != "")
                        {
                            strErrorMessage = strErrorMessage.Substring(0, strErrorMessage.Length - 1);
                            strErrorMessageAll = strErrorMessageAll + strFileList[intI] + "\r\n" + strErrorMessage;
                            continue;
                        }
                        else
                        {
                            m_BackupLogPath = m_strFolder + @"\Backup";
                            if (this.m_strOUTFolder == "")
                            {
                                if (m_BackupLogPath.Trim() == "") { m_BackupLogPath = @Directory.GetCurrentDirectory().Trim() + @"\BackupLog"; }
                            }
                            if (!Directory.Exists(m_BackupLogPath)) { Directory.CreateDirectory(m_BackupLogPath); }
                            File.Move(@m_strFolder + @"\" + strFileList[intI], m_BackupLogPath + @"\" + strFileList[intI]);
                        }
                    }
                }
            }
            saveLog(strErrorMessageAll);
        }
        #endregion
    }
}
