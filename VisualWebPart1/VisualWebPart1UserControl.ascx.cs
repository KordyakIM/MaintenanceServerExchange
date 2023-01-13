using System;
using System.Web.UI;
using System.Management.Automation;
using System.Text;
using Microsoft.SharePoint;
using System.Web.UI.WebControls;
using System.IO;
using System.Web;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Data;
using System.Collections.ObjectModel;
using System.Management.Automation.Runspaces;
using System.Diagnostics;

namespace MaintenanceServers2019.VisualWebPart1
{
    public partial class VisualWebPart1UserControl : UserControl
    {
        string NameMailboxServer;
        string ClusterNodeStatus;
        string DBCopyAutoActivationPolicy;
        string DBCopyActivationDisabledAndMoveNow;
        string DBPreference;
        string MountedDatabase;
        string QueueNumber;
        string MailboxNumber;
        string TimeStamp;
        string script;
        string path;
        string Name_Server;
        PowerShell ps;
        DataRow dr; //add to row
        DataTable dt;
        bool check_comboBox;
        bool result_upper;
        bool result_lower;
        bool bool_check_email;
        int number_comboBox;


        void Create_GridTable()
        {
            dt = new DataTable(); //create table
            ////Добавление колонок в DataTable переменную
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("MaintenanceStatus", typeof(string));
            dt.Columns.Add("Ping", typeof(string));
            dt.Columns.Add("ClusterNodeStatus", typeof(string));
            dt.Columns.Add("DBCopyAutoActivationPolicy", typeof(string));
            dt.Columns.Add("DBCopyActivationDisabledAndMoveNow", typeof(string));
            dt.Columns.Add("DBPreference", typeof(string));
            dt.Columns.Add("MountedDatabase", typeof(string));
            dt.Columns.Add("QueueNumber", typeof(string));
            dt.Columns.Add("MailboxNumber", typeof(string));
            dt.Columns.Add("TimeStamp", typeof(string));
        }
        void Add_to_GridTable(string grid_name, string grid_MaintenanceStatus, string grid_Ping, string grid_ClusterNodeStatus, string grid_DatabaseCopyAutoActivationPolicy, string grid_DatabaseCopyActivationDisabledAndMoveNow, string grid_DBPreference, string grid_MountedDatabase, string grid_QueueNumber, string grid_MailboxNumber, string grid_TimeStamp)
        {
            dr = dt.NewRow();
            dr["Name"] = grid_name;
            dr["MaintenanceStatus"] = grid_MaintenanceStatus;
            dr["Ping"] = grid_Ping;
            dr["ClusterNodeStatus"] = grid_ClusterNodeStatus;
            dr["DBCopyAutoActivationPolicy"] = grid_DatabaseCopyAutoActivationPolicy;
            dr["DBCopyActivationDisabledAndMoveNow"] = grid_DatabaseCopyActivationDisabledAndMoveNow;
            dr["DBPreference"] = grid_DBPreference;
            dr["MountedDatabase"] = grid_MountedDatabase;
            dr["QueueNumber"] = grid_QueueNumber;
            dr["MailboxNumber"] = grid_MailboxNumber;
            dr["TimeStamp"] = grid_TimeStamp;
            dt.Rows.Add(dr);
        }
        void Script_Exchange()
        {
            script = @"
            $password = ConvertTo-SecureString 'Password' -AsPlainText -Force;
            $cred = New-Object System.Management.Automation.PSCredential('Username', $password);
            $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck;
            $RemoteExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://owa.ru/PowerShell/ -Credential $cred -Authentication Basic -SessionOption $SessionOptions -ErrorAction:SilentlyContinue; 
            Import-PSSession $RemoteExSession -AllowClobber -CommandName Get-Mailboxserver, Get-MailboxDatabaseCopyStatus, Get-Queue, Get-Mailbox, Get-MailboxDatabase;
            $MailboxServers = Get-Mailboxserver | ?{$_.name -like 'se*'} | sort name | select name,DatabaseCopyAutoActivationPolicy,DatabaseCopyActivationDisabledAndMoveNow;
            $AllDB_true = (Get-MailboxDatabase | ?{$_.server -like 'se*'}).count
            $array = @() 
            $AllMB = 0
            $AllDB = 0
            $AllQueue = 0
            $chekDBPreference_Boolean = $false
            foreach ($MailboxServer in $MailboxServers){
                $Path = 'C:\folder\'+$MailboxServer.name+'.csv'
                $Exchange_array = @(import-csv $Path)
                $ClusterNodeStatus = (Invoke-Command -ComputerName $MailboxServer.name -Credential $cred {return Get-ClusterNode -Name $env:COMPUTERNAME | select -ExpandProperty state}).value
                if(!$ClusterNodeStatus){$ClusterNodeStatus = 'Error'}
                $DBPreference = (Get-MailboxDatabaseCopyStatus -Server $MailboxServer.name | ?{$_.ActivationPreference -eq '1'} | select -ExpandProperty databasename).count
                if($DBPreference -ne $Exchange_array.MountedDatabase){$chekDBPreference_Boolean = $true}
                $Object = New-Object PSObject
                $Object | add-member Noteproperty -Name 'Name' -Value $MailboxServer.name -Force
                $Object | add-member Noteproperty -Name 'ClusterNodeStatus' -Value $ClusterNodeStatus -Force
                $Object | add-member Noteproperty -Name 'DBCopyAutoActivationPolicy' -Value $MailboxServer.DatabaseCopyAutoActivationPolicy -Force
                $Object | add-member Noteproperty -Name 'DBCopyActivationDisabledAndMoveNow' -Value $MailboxServer.DatabaseCopyActivationDisabledAndMoveNow -Force
                $Object | add-member Noteproperty -Name 'DBPreference' -Value $DBPreference -Force
                $Object | add-member Noteproperty -Name 'MountedDatabase' -Value $Exchange_array.MountedDatabase -Force
                $Object | add-member Noteproperty -Name 'QueueNumber' -Value $Exchange_array.QueueNumber -Force
                $Object | add-member Noteproperty -Name 'MailboxNumber' -Value $Exchange_array.MailboxNumber -Force
                $Object | add-member Noteproperty -Name 'TimeStamp' -Value $Exchange_array.TimeStamp -Force
                $array += $Object
                $AllMB += $Exchange_array.MailboxNumber
                $AllDB += $Exchange_array.MountedDatabase
                $AllQueue += $Exchange_array.QueueNumber
            }
            
            $Object = New-Object PSObject
            $Object | add-member Noteproperty -Name 'Name' -Value 'All' -Force
            $Object | add-member Noteproperty -Name 'ClusterNodeStatus' -Value '' -Force
            $Object | add-member Noteproperty -Name 'DBCopyAutoActivationPolicy' -Value '' -Force
            $Object | add-member Noteproperty -Name 'DBCopyActivationDisabledAndMoveNow' -Value '' -Force
            if($chekDBPreference_Boolean){
                $Object | add-member Noteproperty -Name 'DBPreference' -Value 'Unmatched' -Force
            }else{
                $Object | add-member Noteproperty -Name 'DBPreference' -Value '' -Force
            }
            if($AllDB -eq $AllDB_true){
                $Object | add-member Noteproperty -Name 'MountedDatabase' -Value $AllDB -Force
                $Object | add-member Noteproperty -Name 'QueueNumber' -Value $AllQueue -Force
                $Object | add-member Noteproperty -Name 'MailboxNumber' -Value $AllMB -Force
            }else{
                $Object | add-member Noteproperty -Name 'MountedDatabase' -Value 'Unmatched' -Force
                $Object | add-member Noteproperty -Name 'QueueNumber' -Value 'Unmatched' -Force
                $Object | add-member Noteproperty -Name 'MailboxNumber' -Value 'Unmatched' -Force
            }
            $Object | add-member Noteproperty -Name 'TimeStamp' -Value '' -Force
            $array += $Object
            $array";
        }
        void Script_Exchange_Reload()
        {
            string target_server = comboBox.SelectedItem.Text;
            script = @"
            $password = ConvertTo-SecureString 'Password' -AsPlainText -Force;
            $cred = New-Object System.Management.Automation.PSCredential('Username', $password);
            $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck;
            $RemoteExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://" + target_server + @".ru/PowerShell/ -Credential $cred -Authentication Basic -SessionOption $SessionOptions -ErrorAction:SilentlyContinue; 
            Import-PSSession $RemoteExSession -AllowClobber -CommandName Get-Mailboxserver, Get-MailboxDatabaseCopyStatus, Get-Queue, Get-Mailbox, Get-MailboxDatabase;
            $MailboxServers = Get-Mailboxserver | ?{$_.name -like 'se*'} | sort name | select name,DatabaseCopyAutoActivationPolicy,DatabaseCopyActivationDisabledAndMoveNow;
            $AllDB_true = (Get-MailboxDatabase | ?{$_.server -like 'se*'}).count
            $array = @() 
            $AllMB = 0
            $AllDB = 0
            $AllQueue = 0
            $chekDBPreference_Boolean = $false
            foreach ($MailboxServer in $MailboxServers){
                $Path = 'C:\Folder\'+$MailboxServer.name+'.csv'
                $Exchange_array = @(import-csv $Path)
                $ClusterNodeStatus = (Invoke-Command -ComputerName $MailboxServer.name -Credential $cred {return Get-ClusterNode -Name $env:COMPUTERNAME | select -ExpandProperty state}).value
                if(!$ClusterNodeStatus){$ClusterNodeStatus = 'Error'}
                $DBPreference = (Get-MailboxDatabaseCopyStatus -Server $MailboxServer.name | ?{$_.ActivationPreference -eq '1'} | select -ExpandProperty databasename).count
                if($DBPreference -ne $Exchange_array.MountedDatabase){$chekDBPreference_Boolean = $true}
                $Object = New-Object PSObject
                $Object | add-member Noteproperty -Name 'Name' -Value $MailboxServer.name -Force
                $Object | add-member Noteproperty -Name 'ClusterNodeStatus' -Value $ClusterNodeStatus -Force
                $Object | add-member Noteproperty -Name 'DBCopyAutoActivationPolicy' -Value $MailboxServer.DatabaseCopyAutoActivationPolicy -Force
                $Object | add-member Noteproperty -Name 'DBCopyActivationDisabledAndMoveNow' -Value $MailboxServer.DatabaseCopyActivationDisabledAndMoveNow -Force
                $Object | add-member Noteproperty -Name 'DBPreference' -Value $DBPreference -Force
                $Object | add-member Noteproperty -Name 'MountedDatabase' -Value $Exchange_array.MountedDatabase -Force
                $Object | add-member Noteproperty -Name 'QueueNumber' -Value $Exchange_array.QueueNumber -Force
                $Object | add-member Noteproperty -Name 'MailboxNumber' -Value $Exchange_array.MailboxNumber -Force
                $Object | add-member Noteproperty -Name 'TimeStamp' -Value $Exchange_array.TimeStamp -Force
                $array += $Object
                $AllMB += $Exchange_array.MailboxNumber
                $AllDB += $Exchange_array.MountedDatabase
                $AllQueue += $Exchange_array.QueueNumber
            }
            
            $Object = New-Object PSObject
            $Object | add-member Noteproperty -Name 'Name' -Value 'All' -Force
            $Object | add-member Noteproperty -Name 'ClusterNodeStatus' -Value '' -Force
            $Object | add-member Noteproperty -Name 'DBCopyAutoActivationPolicy' -Value '' -Force
            $Object | add-member Noteproperty -Name 'DBCopyActivationDisabledAndMoveNow' -Value '' -Force
            if($chekDBPreference_Boolean){
                $Object | add-member Noteproperty -Name 'DBPreference' -Value 'Unmatched' -Force
            }else{
                $Object | add-member Noteproperty -Name 'DBPreference' -Value '' -Force
            }
            if($AllDB -eq $AllDB_true){
                $Object | add-member Noteproperty -Name 'MountedDatabase' -Value $AllDB -Force
                $Object | add-member Noteproperty -Name 'QueueNumber' -Value $AllQueue -Force
                $Object | add-member Noteproperty -Name 'MailboxNumber' -Value $AllMB -Force
            }else{
                $Object | add-member Noteproperty -Name 'MountedDatabase' -Value 'Unmatched' -Force
                $Object | add-member Noteproperty -Name 'QueueNumber' -Value 'Unmatched' -Force
                $Object | add-member Noteproperty -Name 'MailboxNumber' -Value 'Unmatched' -Force
            }
            $Object | add-member Noteproperty -Name 'TimeStamp' -Value '' -Force
            $array += $Object
            $array";
        }
        void get_value_outp(PSObject outp)
        {
            try { NameMailboxServer = Convert.ToString(outp.Properties["Name"].Value); } catch (Exception) { NameMailboxServer = "Null"; }
            try { ClusterNodeStatus = Convert.ToString(outp.Properties["ClusterNodeStatus"].Value); } catch (Exception) { ClusterNodeStatus = "Null"; }
            try { DBCopyAutoActivationPolicy = Convert.ToString(outp.Properties["DBCopyAutoActivationPolicy"].Value); } catch (Exception) { DBCopyAutoActivationPolicy = "Null"; }
            try { DBCopyActivationDisabledAndMoveNow = Convert.ToString(outp.Properties["DBCopyActivationDisabledAndMoveNow"].Value); } catch (Exception) { DBCopyActivationDisabledAndMoveNow = "Null"; }
            try { DBPreference = Convert.ToString(outp.Properties["DBPreference"].Value); } catch (Exception) { DBPreference = "Null"; }
            try { MountedDatabase = Convert.ToString(outp.Properties["MountedDatabase"].Value); } catch (Exception) { MountedDatabase = "Null"; }
            try { QueueNumber = Convert.ToString(outp.Properties["QueueNumber"].Value); } catch (Exception) { QueueNumber = "Null"; }
            try { MailboxNumber = Convert.ToString(outp.Properties["MailboxNumber"].Value); } catch (Exception) { MailboxNumber = "Null"; }
            try { TimeStamp = Convert.ToString(outp.Properties["TimeStamp"].Value); } catch (Exception) { TimeStamp = "Null"; }
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            check_comboBox = false;
            path = @"C:\Logs\Maintenance.log";
            Create_GridTable();
            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();
            bool_check_email = false;

            Page.Server.ScriptTimeout = 6400; // specify the timeout to 3600 seconds
            if (!IsPostBack)
            {
                using (SPLongOperation operation = new SPLongOperation(Page))
                {
                    Script_Exchange();
                    using (ps = PowerShell.Create())
                    {
                        PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                        ps.AddScript(script);
                        IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                        ps.EndInvoke(result);
                        ps.Stop();
                        ResultBox.Text = string.Empty;
                        string check = string.Empty;
                        //if (output.) {
                            for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                            {
                                foreach (PSObject outp in output)
                                {
                                    get_value_outp(outp);
                                    if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.ToString().Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToUpper();
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                    }else if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) &!check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                    }
                                }
                                result_upper = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                                if (result_upper)
                                {
                                    comboBox.Items.Add(new ListItem(CheckBoxList_array01.Items[i].Text, CheckBoxList_array01.Items[i].Text));
                                }
                            }

                            for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                            {
                                foreach (PSObject outp in output)
                                {
                                    get_value_outp(outp);
                                    if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToUpper();
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                    }
                                    else if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                    }
                                }
                                result_upper = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                                if (result_upper)
                                {
                                    comboBox.Items.Add(new ListItem(CheckBoxList_array02.Items[i].Text, CheckBoxList_array02.Items[i].Text));
                                }
                            }
                        //}
                        foreach (PSObject outp in output)
                        {
                            get_value_outp(outp);
                            if (string.Equals("All", NameMailboxServer, StringComparison.CurrentCultureIgnoreCase))
                            {
                                Add_to_GridTable(NameMailboxServer, "", "", ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                            }
                        }
                    }
                    comboBox.DataBind();
                    GridView1.DataSource = dt;
                    GridView1.DataBind();
                }
            }
        }
        protected void Start_click(object sender, EventArgs e)
        {
            //chenge item in comboBox
            check_comboBox = false;
            for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
            {
                result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                if (!CheckBoxList_array01.Items[i].Selected & !check_comboBox & result_lower)
                {
                    check_comboBox = true;
                    comboBox.SelectedValue = CheckBoxList_array01.Items[i].Text;
                }
            }
            for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
            {
                result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                if (!CheckBoxList_array02.Items[i].Selected & !check_comboBox & result_lower)
                {
                    check_comboBox = true;
                    comboBox.SelectedValue = CheckBoxList_array02.Items[i].Text;
                }
            }
            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();

            Page.Server.ScriptTimeout = 6400; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ResultBox.Text = string.Empty;
                if (TextBox1.Text == "Exchange")
                {
                    if (!String.IsNullOrEmpty(TextBox_Goal.Text))
                    {
                        string target_server = comboBox.SelectedItem.Text;
                        string check = string.Empty;
                        for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                        {
                            if (CheckBoxList_array01.Items[i].Selected)
                            {
                                check += CheckBoxList_array01.Items[i].Text;
                            }
                        }
                        for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                        {
                            if (CheckBoxList_array02.Items[i].Selected)
                            {
                                check += CheckBoxList_array02.Items[i].Text;
                            }
                        }
                        if (System.Text.RegularExpressions.Regex.IsMatch(check, "sn", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        {
                            if (!check.Contains(target_server))
                            {
                                result_lower = Char.IsLower(target_server, 2);
                                if (result_lower)
                                {
                                    comboBox.Items.Clear(); //clearing combobox
                                    for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                    {
                                        if (CheckBoxList_array01.Items[i].Selected)
                                        {
                                            //check life server through ping
                                            PingReply reply = pinger.Send(CheckBoxList_array01.Items[i].Text);
                                            pingable = reply.Status == IPStatus.Success;
                                            if (pingable)
                                            {
                                                script = @"C:\folder\Start-ExchangeServerMaintenanceMode.ps1 -SourceServer '" + CheckBoxList_array01.Items[i].Text + "' -TargetServerFQDN '" + target_server + ".npr.nornick.ru' -UserName " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + " -Goal '" + TextBox_Goal.Text + "'";
                                                using (ps = PowerShell.Create())
                                                {
                                                    var builder = new StringBuilder();
                                                    PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                    ps.AddScript(script);
                                                    IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                    ps.EndInvoke(invocation);
                                                    ps.Stop();
                                                    foreach (PSObject outp in output)
                                                    {
                                                        if (!outp.BaseObject.ToString().Contains("tmp"))
                                                        {
                                                            builder.Append(outp.BaseObject.ToString() + "\r\n");
                                                            if (outp.BaseObject.ToString().Contains("SUCCESS: Done! Server " + CheckBoxList_array01.Items[i].Text + " is put succesfully into maintenance mode"))
                                                            {
                                                                this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToUpper();
                                                                Name_Server = CheckBoxList_array01.Items[i].Text.ToUpper();
                                                            }
                                                        }
                                                    }
                                                    ResultBox.Text += builder.ToString();
                                                }
                                            }
                                            else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_array01.Items[i].Text; }
                                        }
                                        result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                                        if (result_lower)
                                        {
                                            comboBox.Items.Add(new ListItem(CheckBoxList_array01.Items[i].Text, CheckBoxList_array01.Items[i].Text));
                                        }
                                    }
                                    for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                    {
                                        if (CheckBoxList_array02.Items[i].Selected)
                                        {
                                            //check life server through ping
                                            PingReply reply = pinger.Send(CheckBoxList_array02.Items[i].Text);
                                            pingable = reply.Status == IPStatus.Success;
                                            if (pingable)
                                            {
                                                script = @"C:\folder\Start-ExchangeServerMaintenanceMode.ps1 -SourceServer '" + CheckBoxList_array02.Items[i].Text + "' -TargetServerFQDN '" + target_server + ".npr.nornick.ru' -UserName " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + " -Goal '" + TextBox_Goal.Text + "'";
                                                using (ps = PowerShell.Create())
                                                {
                                                    var builder = new StringBuilder();
                                                    PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                    ps.AddScript(script);
                                                    IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                    ps.EndInvoke(invocation);
                                                    ps.Stop();
                                                    foreach (PSObject outp in output)
                                                    {
                                                        if (!outp.BaseObject.ToString().Contains("tmp"))
                                                        {
                                                            builder.Append(outp.BaseObject.ToString() + "\r\n");
                                                            if (outp.BaseObject.ToString().Contains("SUCCESS: Done! Server " + CheckBoxList_array02.Items[i].Text + " is put succesfully into maintenance mode"))
                                                            {
                                                                this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToUpper();
                                                                Name_Server = CheckBoxList_array02.Items[i].Text.ToUpper();
                                                            }
                                                        }
                                                    }
                                                    ResultBox.Text += builder.ToString();
                                                }
                                            }
                                            else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_array02.Items[i].Text; }
                                        }
                                        result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                                        if (result_lower)
                                        {
                                            comboBox.Items.Add(new ListItem(CheckBoxList_array02.Items[i].Text, CheckBoxList_array02.Items[i].Text));
                                        }
                                    }
                                    comboBox.DataBind();
                                    File.AppendAllText(path, ResultBox.Text);

                                    //update inform to grid-----------------------------------------------------------------------------------------
                                    Create_GridTable();
                                    Script_Exchange_Reload();
                                    //Script_Exchange();
                                    using (ps = PowerShell.Create())
                                    {
                                        PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                        ps.AddScript(script);
                                        IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                        ps.EndInvoke(result);
                                        ps.Stop();
                                        check = string.Empty;
                                        for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                        {
                                            foreach (PSObject outp in output)
                                            {
                                                get_value_outp(outp);
                                                if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                                {
                                                    this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToUpper();
                                                    //check life server through ping
                                                    PingReply reply = pinger.Send(NameMailboxServer);
                                                    pingable = reply.Status == IPStatus.Success;
                                                    check += NameMailboxServer;
                                                    Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                                }
                                                else if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                                {
                                                    //check life server through ping
                                                    PingReply reply = pinger.Send(NameMailboxServer);
                                                    pingable = reply.Status == IPStatus.Success;
                                                    check += NameMailboxServer;
                                                    Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                                }
                                            }
                                        }

                                        for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                        {
                                            foreach (PSObject outp in output)
                                            {
                                                get_value_outp(outp);
                                                if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                                {
                                                    this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToUpper();
                                                    //check life server through ping
                                                    PingReply reply = pinger.Send(NameMailboxServer);
                                                    pingable = reply.Status == IPStatus.Success;
                                                    check += NameMailboxServer;
                                                    Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                                }
                                                else if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                                {
                                                    //check life server through ping
                                                    PingReply reply = pinger.Send(NameMailboxServer);
                                                    pingable = reply.Status == IPStatus.Success;
                                                    check += NameMailboxServer;
                                                    Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                                }
                                            }
                                        }
                                        foreach (PSObject outp in output)
                                        {
                                            get_value_outp(outp);
                                            if (string.Equals("All", NameMailboxServer, StringComparison.CurrentCultureIgnoreCase))
                                            {
                                                Add_to_GridTable(NameMailboxServer, "", "", ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                            }
                                        }
                                    }
                                    GridView1.DataSource = dt;
                                    GridView1.DataBind();
                                    //update inform to grid-----------------------------------------------------------------------------------------
                                    //Send email
                                    for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                    {
                                        if (CheckBoxList_array01.Items[i].Selected)
                                        {
                                            result_upper = Char.IsUpper(CheckBoxList_array01.Items[i].Text, 2);
                                            if (result_upper)
                                            {
                                                bool_check_email = true;
                                            }
                                        }
                                    }
                                    for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                    {
                                        if (CheckBoxList_array02.Items[i].Selected)
                                        {
                                            result_upper = Char.IsUpper(CheckBoxList_array02.Items[i].Text, 2);
                                            if (result_upper)
                                            {
                                                bool_check_email = true;
                                            }
                                        }
                                    }
                                    if (bool_check_email) {email(ResultBox.Text, "Start Maintenance", Name_Server);}
                                }
                                else { ResultBox.Text = "Target Server - " + target_server + " is already exists in Maintenance! Choose another Server!"; }
                            }
                            else { ResultBox.Text = "Target Server - " + target_server + " is already exists in Source Server! Choose another Server!"; }
                        }
                        else { ResultBox.Text = "Choose checkbox, please."; }
                    }
                    else { ResultBox.Text = "Input you goal, please."; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void Stop_click(object sender, EventArgs e)
        {
            //chenge item in comboBox
            check_comboBox = false;
            for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
            {
                result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                if (!CheckBoxList_array01.Items[i].Selected & !check_comboBox & result_lower)
                {
                    check_comboBox = true;
                    comboBox.SelectedValue = CheckBoxList_array01.Items[i].Text;
                }
            }
            for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
            {
                result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                if (!CheckBoxList_array02.Items[i].Selected & !check_comboBox & result_lower)
                {
                    check_comboBox = true;
                    comboBox.SelectedValue = CheckBoxList_array02.Items[i].Text;
                }
            }

            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();

            Page.Server.ScriptTimeout = 6400; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                ResultBox.Text = string.Empty;
                if (TextBox1.Text == "Exchange")
                {
                    if (!String.IsNullOrEmpty(TextBox_Goal.Text))
                    {
                        string check = string.Empty;
                        for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                        {
                            if (CheckBoxList_array01.Items[i].Selected)
                            {
                                check += CheckBoxList_array01.Items[i].Text;
                            }
                        }
                        for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                        {
                            if (CheckBoxList_array02.Items[i].Selected)
                            {
                                check += CheckBoxList_array02.Items[i].Text;
                            }
                        }
                        if (System.Text.RegularExpressions.Regex.IsMatch(check, "sn", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        {
                            result_lower = Char.IsLower(comboBox.SelectedItem.Text, 2);
                            if (result_lower)
                            {
                                comboBox.Items.Clear(); //clearing combobox
                                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                {
                                    if (CheckBoxList_array01.Items[i].Selected)
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(CheckBoxList_array01.Items[i].Text);
                                        pingable = reply.Status == IPStatus.Success;
                                        if (pingable)
                                        {
                                            script = @"C:\folder\Stop-ExchangeServerMaintenanceMode.ps1 -Server '" + CheckBoxList_array01.Items[i].Text + "' -UserName " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + " -Goal '" + TextBox_Goal.Text + "'";
                                            using (ps = PowerShell.Create())
                                            {
                                                var builder = new StringBuilder();
                                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                ps.AddScript(script);
                                                IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                ps.EndInvoke(invocation);
                                                ps.Stop();
                                                foreach (PSObject outp in output)
                                                {
                                                    if (!outp.BaseObject.ToString().Contains("tmp"))
                                                    {
                                                        builder.Append(outp.BaseObject.ToString() + "\r\n");
                                                        if (string.Equals(outp.BaseObject.ToString(), "SUCCESS: Done! Server " + CheckBoxList_array01.Items[i].Text + " successfully taken out of Maintenance Mode.", StringComparison.CurrentCultureIgnoreCase))
                                                        {
                                                            this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToLower();
                                                            Name_Server = CheckBoxList_array01.Items[i].Text.ToLower();
                                                        }
                                                    }
                                                }
                                                ResultBox.Text += builder.ToString();
                                            }
                                        }
                                        else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_array01.Items[i].Text; }
                                    }
                                    result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                                    if (result_lower)
                                    {
                                        comboBox.Items.Add(new ListItem(CheckBoxList_array01.Items[i].Text, CheckBoxList_array01.Items[i].Text));
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                {
                                    if (CheckBoxList_array02.Items[i].Selected)
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(CheckBoxList_array02.Items[i].Text);
                                        pingable = reply.Status == IPStatus.Success;
                                        if (pingable)
                                        {
                                            script = @"C:\folder\Stop-ExchangeServerMaintenanceMode.ps1 -Server '" + CheckBoxList_array02.Items[i].Text + "' -UserName " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + " -Goal '" + TextBox_Goal.Text + "'";
                                            using (ps = PowerShell.Create())
                                            {
                                                var builder = new StringBuilder();
                                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                ps.AddScript(script);
                                                IAsyncResult invocation = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                ps.EndInvoke(invocation);
                                                ps.Stop();
                                                foreach (PSObject outp in output)
                                                {
                                                    if (!outp.BaseObject.ToString().Contains("tmp"))
                                                    {
                                                        builder.Append(outp.BaseObject.ToString() + "\r\n");
                                                        if (string.Equals(outp.BaseObject.ToString(), "SUCCESS: Done! Server " + CheckBoxList_array02.Items[i].Text + " successfully taken out of Maintenance Mode.", StringComparison.CurrentCultureIgnoreCase))
                                                        {
                                                            this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToLower();
                                                            Name_Server = CheckBoxList_array02.Items[i].Text.ToLower();
                                                        }
                                                    }
                                                }
                                                ResultBox.Text += builder.ToString();
                                            }
                                        }
                                        else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_array02.Items[i].Text; }
                                    }
                                    result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                                    if (result_lower)
                                    {
                                        comboBox.Items.Add(new ListItem(CheckBoxList_array02.Items[i].Text, CheckBoxList_array02.Items[i].Text));
                                    }
                                }
                                comboBox.DataBind();
                                File.AppendAllText(path, ResultBox.Text);


                                //update inform to grid-----------------------------------------------------------------------------------------
                                Create_GridTable();
                                Script_Exchange_Reload();
                                using (ps = PowerShell.Create())
                                {
                                    PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                    ps.AddScript(script);
                                    IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                    ps.EndInvoke(result);
                                    ps.Stop();
                                    check = string.Empty;
                                    for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                    {
                                        foreach (PSObject outp in output)
                                        {
                                            get_value_outp(outp);
                                            if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                            {
                                                //check life server through ping
                                                PingReply reply = pinger.Send(NameMailboxServer);
                                                pingable = reply.Status == IPStatus.Success;
                                                check += NameMailboxServer;
                                                Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                            }
                                            else if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                            {
                                                this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToLower();
                                                //check life server through ping
                                                PingReply reply = pinger.Send(NameMailboxServer);
                                                pingable = reply.Status == IPStatus.Success;
                                                check += NameMailboxServer;
                                                Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                            }
                                        }
                                    }

                                    for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                    {
                                        foreach (PSObject outp in output)
                                        {
                                            get_value_outp(outp);
                                            if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                            {
                                                //check life server through ping
                                                PingReply reply = pinger.Send(NameMailboxServer);
                                                pingable = reply.Status == IPStatus.Success;
                                                check += NameMailboxServer;
                                                Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                            }
                                            else if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                            {
                                                this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToLower();
                                                //check life server through ping
                                                PingReply reply = pinger.Send(NameMailboxServer);
                                                pingable = reply.Status == IPStatus.Success;
                                                check += NameMailboxServer;
                                                Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                            }
                                        }
                                    }
                                    foreach (PSObject outp in output)
                                    {
                                        get_value_outp(outp);
                                        if (string.Equals("All", NameMailboxServer, StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            Add_to_GridTable(NameMailboxServer, "", "", ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                        }
                                    }
                                }
                                GridView1.DataSource = dt;
                                GridView1.DataBind();
                                //update inform to grid-----------------------------------------------------------------------------------------
                                //Send email
                                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                                {
                                    if (CheckBoxList_array01.Items[i].Selected)
                                    {
                                        result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                                        if (result_lower) {
                                            bool_check_email = true;
                                        }
                                    }
                                }
                                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                                {
                                    if (CheckBoxList_array02.Items[i].Selected)
                                    {
                                        result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                                        if (result_lower)
                                        {
                                            bool_check_email = true;
                                        }
                                    }
                                }
                                if (bool_check_email) { email(ResultBox.Text, "Stop Maintenance", Name_Server); }
                            }
                            else { ResultBox.Text = "Target Server -" + comboBox.SelectedItem.Text + " is in Maintenance! Choose another Server!"; }
                        }
                        else { ResultBox.Text = "Choose checkbox, please."; }
                    }
                    else { ResultBox.Text = "Input your goal, please."; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void Refresh_click(object sender, EventArgs e)
        {
            result_lower = Char.IsLower(comboBox.SelectedItem.Text, 2);
            if (result_lower)
            {
                //chenge item in comboBox
                check_comboBox = false;
                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                {
                    result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                    if (!CheckBoxList_array01.Items[i].Selected & !check_comboBox & result_lower)
                    {
                        check_comboBox = true;
                        comboBox.SelectedValue = CheckBoxList_array01.Items[i].Text;
                    }
                }
                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                {
                    result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                    if (!CheckBoxList_array02.Items[i].Selected & !check_comboBox & result_lower)
                    {
                        check_comboBox = true;
                        comboBox.SelectedValue = CheckBoxList_array02.Items[i].Text;
                    }
                }
                //-----------------------
                ResultBox.Text = string.Empty;
                Create_GridTable();
                //check life server through ping
                bool pingable = false;
                Ping pinger = null;
                pinger = new Ping();
                Page.Server.ScriptTimeout = 6400; // specify the timeout to 3600 seconds
                using (SPLongOperation operation = new SPLongOperation(Page))
                {
                    Script_Exchange_Reload();
                    comboBox.Items.Clear(); //clearing combobox
                    using (ps = PowerShell.Create())
                    {
                        PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                        ps.AddScript(script);
                        IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                        ps.EndInvoke(result);
                        ps.Stop();
                        string check = string.Empty;
                        for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                        {
                            foreach (PSObject outp in output)
                            {
                                get_value_outp(outp);
                                if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.ToString().Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToUpper();
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                }
                                else if (string.Equals(CheckBoxList_array01.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    this.CheckBoxList_array01.Items[i].Text = CheckBoxList_array01.Items[i].Text.ToLower();
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                }
                            }
                            result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                            if (result_lower)
                            {
                                comboBox.Items.Add(new ListItem(CheckBoxList_array01.Items[i].Text, CheckBoxList_array01.Items[i].Text));
                            }
                        }

                        for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                        {
                            foreach (PSObject outp in output)
                            {
                                get_value_outp(outp);
                                if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & (ClusterNodeStatus.ToString().Equals("Paused") || ClusterNodeStatus.ToString().Equals("Error") || DBCopyAutoActivationPolicy.Equals("Blocked")) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToUpper();
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                }
                                else if (string.Equals(CheckBoxList_array02.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    this.CheckBoxList_array02.Items[i].Text = CheckBoxList_array02.Items[i].Text.ToLower();
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                                }
                            }
                            result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                            if (result_lower)
                            {
                                comboBox.Items.Add(new ListItem(CheckBoxList_array02.Items[i].Text, CheckBoxList_array02.Items[i].Text));
                            }
                        }
                        foreach (PSObject outp in output)
                        {
                            get_value_outp(outp);
                            if (string.Equals("All", NameMailboxServer, StringComparison.CurrentCultureIgnoreCase))
                            {
                                Add_to_GridTable(NameMailboxServer, "", "", ClusterNodeStatus, DBCopyAutoActivationPolicy, DBCopyActivationDisabledAndMoveNow, DBPreference, MountedDatabase, QueueNumber, MailboxNumber, TimeStamp);
                            }
                        }
                    }
                    comboBox.DataBind();
                    GridView1.DataSource = dt;
                    GridView1.DataBind();
                }
            }
            else { ResultBox.Text = "Target Server -" + comboBox.SelectedItem.Text + " is in Maintenance! Choose another Server! Target Server needs for connect to Exchange."; }
        }
        protected void Rebuild_click(object sender, EventArgs e)
        {
            bool result_lower = Char.IsLower(comboBox.SelectedItem.Text, 2);
            if (result_lower)
            {
                //chenge item in comboBox
                check_comboBox = false;
                for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                {
                    result_lower = Char.IsLower(CheckBoxList_array01.Items[i].Text, 2);
                    if (!CheckBoxList_array01.Items[i].Selected & !check_comboBox & result_lower)
                    {
                        check_comboBox = true;
                        comboBox.SelectedValue = CheckBoxList_array01.Items[i].Text;
                    }
                }
                for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                {
                    result_lower = Char.IsLower(CheckBoxList_array02.Items[i].Text, 2);
                    if (!CheckBoxList_array02.Items[i].Selected & !check_comboBox & result_lower)
                    {
                        check_comboBox = true;
                        comboBox.SelectedValue = CheckBoxList_array02.Items[i].Text;
                    }
                }
                //check life server through ping
                bool pingable = false;
                Ping pinger = null;
                pinger = new Ping();
                Page.Server.ScriptTimeout = 6400; // specify the timeout to 3600 seconds
                using (SPLongOperation operation = new SPLongOperation(Page))
                {
                    if (TextBox1.Text == "Rebuild")
                    {
                        ResultBox.Text = string.Empty;
                        string check = string.Empty;
                        for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                        {
                            if (CheckBoxList_array01.Items[i].Selected)
                            {
                                check += CheckBoxList_array01.Items[i].Text;
                            }
                        }
                        for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                        {
                            if (CheckBoxList_array02.Items[i].Selected)
                            {
                                check += CheckBoxList_array02.Items[i].Text;
                            }
                        }
                        if (System.Text.RegularExpressions.Regex.IsMatch(check, "vn", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        {
                            for (int i = 0; i < CheckBoxList_array01.Items.Count; i++)
                            {
                                if (CheckBoxList_array01.Items[i].Selected)
                                {
                                    result_upper = Char.IsUpper(CheckBoxList_array01.Items[i].Text, 2);
                                    if (result_upper)
                                    {
                                        PingReply reply = pinger.Send(CheckBoxList_array01.Items[i].Text);
                                        pingable = reply.Status == IPStatus.Success;
                                        if (pingable)
                                        {
                                            Name_Server = CheckBoxList_array01.Items[i].Text;
                                            script = @"C:\folder\RebuildTransportDB.ps1 -Server '" + CheckBoxList_array01.Items[i].Text + "'";
                                            using (ps = PowerShell.Create())
                                            {
                                                var builder = new StringBuilder();
                                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                ps.AddScript(script);
                                                IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                ps.EndInvoke(result);
                                                ps.Stop();
                                                foreach (PSObject outp in output)
                                                {
                                                    if (!outp.BaseObject.ToString().Contains("tmp"))
                                                    {
                                                        builder.Append(outp.BaseObject.ToString() + "\r\n");
                                                    }
                                                }
                                                ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                                ResultBox.Text += builder.ToString();
                                            }
                                        }
                                        else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_array01.Items[i].Text; }
                                    }
                                    else { ResultBox.Text = "Данный сервер не на обслуживании!"; }
                                }
                            }
                            for (int i = 0; i < CheckBoxList_array02.Items.Count; i++)
                            {
                                if (CheckBoxList_array02.Items[i].Selected)
                                {
                                    result_upper = Char.IsUpper(CheckBoxList_array02.Items[i].Text, 2);
                                    if (result_upper)
                                    {
                                        PingReply reply = pinger.Send(CheckBoxList_array02.Items[i].Text);
                                        pingable = reply.Status == IPStatus.Success;
                                        if (pingable)
                                        {
                                            Name_Server = CheckBoxList_array02.Items[i].Text;
                                            script = @"C:\folder\RebuildTransportDB.ps1 -Server '" + CheckBoxList_array02.Items[i].Text + "'";
                                            using (ps = PowerShell.Create())
                                            {
                                                var builder = new StringBuilder();
                                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                                ps.AddScript(script);
                                                IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                                ps.EndInvoke(result);
                                                ps.Stop();
                                                foreach (PSObject outp in output)
                                                {
                                                    if (!outp.BaseObject.ToString().Contains("tmp"))
                                                    {
                                                        builder.Append(outp.BaseObject.ToString() + "\r\n");
                                                    }
                                                }
                                                ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                                ResultBox.Text += builder.ToString();
                                            }
                                        }
                                        else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_array02.Items[i].Text; }
                                    }
                                    else { ResultBox.Text = "Данный сервер не на обслуживании!"; }
                                }
                            }
                            email(ResultBox.Text, "Rebuild Transport DB", Name_Server); //отправка сообщения
                            File.AppendAllText(path, ResultBox.Text);
                        }
                        else { ResultBox.Text = "Choose checkbox, please!"; }
                    }
                    else { ResultBox.Text = "Wrong Input Code!"; }
                }
            }
            else { ResultBox.Text = "Target Server -" + comboBox.SelectedItem.Text + " is in Maintenance! Choose another Server! Target Server needs for connect to Exchange."; }
        }
        void email(string body, string action, string Server)
        {
            try
            {
                // отправитель - устанавливаем адрес и отображаемое в письме имя
                MailAddress from = new MailAddress("SourceAccount@mail.ru", "Maintenance Exchange");
                // кому отправляем
                MailAddress to = new MailAddress("DestinationAccount@mail.ru");
                // создаем объект сообщения
                MailMessage m = new MailMessage(from, to);
                m.Subject = action + " -> " + Server;
                m.Body = body;
                m.To.Add(to_IT);
                // письмо представляет код html
                SmtpClient smtp = new SmtpClient("mail.ru", 25);
                smtp.Send(m);
            }
            catch (Exception e)
            {}
        }
    }
}
