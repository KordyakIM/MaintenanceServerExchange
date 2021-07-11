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
        string DatabaseCopyAutoActivationPolicy;
        string DatabaseCopyActivationDisabledAndMoveNow;
        string MountedDatabase;
        string QueueNumber;
        string MessagesNumber;
        string script;
        PowerShell ps;
        DataRow dr; //add to row
        DataTable dt;
        void Create_GridTable()
        {
            dt = new DataTable(); //create table
            ////Добавление колонок в DataTable переменную
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("MaintenanceStatus", typeof(string));
            dt.Columns.Add("Ping", typeof(string));
            dt.Columns.Add("DatabaseCopyAutoActivationPolicy", typeof(string));
            dt.Columns.Add("DatabaseCopyActivationDisabledAndMoveNow", typeof(string));
            dt.Columns.Add("MountedDatabase", typeof(string));
            dt.Columns.Add("QueueNumber", typeof(string));
            dt.Columns.Add("MessagesNumber", typeof(string));
        }
        void Add_to_GridTable(string grid_name, string grid_MaintenanceStatus, string grid_Ping, string grid_DatabaseCopyAutoActivationPolicy, string grid_DatabaseCopyActivationDisabledAndMoveNow, string grid_MountedDatabase, string grid_QueueNumber, string grid_MessagesNumber)
        {
            dr = dt.NewRow();
            dr["Name"] = grid_name;
            dr["MaintenanceStatus"] = grid_MaintenanceStatus;
            dr["Ping"] = grid_Ping;
            dr["DatabaseCopyAutoActivationPolicy"] = grid_DatabaseCopyAutoActivationPolicy;
            dr["DatabaseCopyActivationDisabledAndMoveNow"] = grid_DatabaseCopyActivationDisabledAndMoveNow;
            dr["MountedDatabase"] = grid_MountedDatabase;
            dr["QueueNumber"] = grid_QueueNumber;
            dr["MessagesNumber"] = grid_MessagesNumber;
            dt.Rows.Add(dr);
        }
        void Script_Exchange()
        {
            script = @"
                    cls
                    $password = ConvertTo-SecureString 'Password' -AsPlainText -Force
                    $cred = New-Object System.Management.Automation.PSCredential('username', $password)
                    if (-not($RemoteEx2013Session)){
                        $SessionOptions = New-PSSessionOption –SkipCACheck –SkipCNCheck –SkipRevocationCheck
                        $RemoteEx2013Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://server/PowerShell/ -Authentication Basic -Credential $cred -SessionOption $SessionOptions 
                        Import-PSSession $RemoteEx2013Session -AllowClobber -CommandName Get-Mailboxserver, Get-MailboxDatabaseCopyStatus, Get-Queue, Get-Mailbox
                    }
                    $MailboxServers = Get-Mailboxserver | ?{$_.name -like 'server'} | sort name | select name,DatabaseCopyAutoActivationPolicy,DatabaseCopyActivationDisabledAndMoveNow 
                    foreach($MailboxServer in $MailboxServers){
                        $MessagesNumber=0
                        $Databasename = Get-MailboxDatabaseCopyStatus -Server $MailboxServer.name | ?{$_.status -eq 'Mounted'} | select databasename
                        $MountedDatabase = ($Databasename | Measure-Object).count
                        $Databasename | %{$MessagesNumber+=(Get-Mailbox -Database $_.databasename).count}	                    
                        #$MountedDatabase = (Get-MailboxDatabaseCopyStatus -Server $MailboxServer.name | ?{$_.status -eq 'Mounted'} | Measure-Object).count
                        $QueueNumber = (Get-Queue -Server $MailboxServer.name | ?{$_.Identity -notlike '*Poison*' -and $_.Identity -notlike '*Shadow*' -and $_.MessageCount -ne 0 -and $_.Status -eq 'Ready'} | select -ExpandProperty MessageCount | Measure-Object).count
	                    $Object = New-Object PSObject
	                    $Object | add-member Noteproperty -Name 'Name' -Value $MailboxServer.name -Force
	                    $Object | add-member Noteproperty -Name 'DatabaseCopyAutoActivationPolicy' -Value $MailboxServer.DatabaseCopyAutoActivationPolicy -Force
	                    $Object | add-member Noteproperty -Name 'DatabaseCopyActivationDisabledAndMoveNow' -Value $MailboxServer.DatabaseCopyActivationDisabledAndMoveNow -Force
	                    $Object | add-member Noteproperty -Name 'MountedDatabase' -Value $MountedDatabase -Force
                        $Object | add-member Noteproperty -Name 'QueueNumber' -Value $QueueNumber -Force
                        $Object | add-member Noteproperty -Name 'MessagesNumber' -Value $MessagesNumber -Force
	                    $Object
                    }";
        }
        void get_value_outp(PSObject outp)
        {
            try { NameMailboxServer = Convert.ToString(outp.Properties["Name"].Value); } catch (Exception) { NameMailboxServer = "Null"; }
            try { DatabaseCopyAutoActivationPolicy = Convert.ToString(outp.Properties["DatabaseCopyAutoActivationPolicy"].Value); } catch (Exception) { DatabaseCopyAutoActivationPolicy = "Null"; }
            try { DatabaseCopyActivationDisabledAndMoveNow = Convert.ToString(outp.Properties["DatabaseCopyActivationDisabledAndMoveNow"].Value); } catch (Exception) { DatabaseCopyActivationDisabledAndMoveNow = "Null"; }
            try { MountedDatabase = Convert.ToString(outp.Properties["MountedDatabase"].Value); } catch (Exception) { MountedDatabase = "Null"; }
            try { QueueNumber = Convert.ToString(outp.Properties["QueueNumber"].Value); } catch (Exception) { QueueNumber = "Null"; }
            try { MessagesNumber = Convert.ToString(outp.Properties["MessagesNumber"].Value); } catch (Exception) { MessagesNumber = "Null"; }
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            Create_GridTable();
            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();

            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
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
                            for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                            {
                                foreach (PSObject outp in output)
                                {
                                    get_value_outp(outp);
                                    if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.ToString().Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        this.CheckBoxList_Norilsk.Items[i].Text = CheckBoxList_Norilsk.Items[i].Text.ToUpper();
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                    }else if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) &!check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                    }
                                }
                                comboBox.Items.Add(new ListItem(CheckBoxList_Norilsk.Items[i].Text, CheckBoxList_Norilsk.Items[i].Text));
                            }

                            for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                            {
                                foreach (PSObject outp in output)
                                {
                                    get_value_outp(outp);
                                    if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        this.CheckBoxList_Talnakh.Items[i].Text = CheckBoxList_Talnakh.Items[i].Text.ToUpper();
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                    }
                                    else if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                    }
                                }
                                comboBox.Items.Add(new ListItem(CheckBoxList_Talnakh.Items[i].Text, CheckBoxList_Talnakh.Items[i].Text));
                            }
                        //}

                    }
                    comboBox.DataBind();
                    GridView1.DataSource = dt;
                    GridView1.DataBind();
                }
            }
        }
        protected void Start_click(object sender, EventArgs e)
        {
            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();

            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                string path = @"C:\work\Exchange\MaintenanceServer\Maintenance.log";
                ResultBox.Text = string.Empty;
                if (TextBox1.Text == "Exchange")
                {
                    string target_server = comboBox.SelectedItem.Text;
                    string check = string.Empty;
                    for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                    {
                        if (CheckBoxList_Norilsk.Items[i].Selected)
                        {
                            check += CheckBoxList_Norilsk.Items[i].Text;
                        }
                    }
                    for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                    {
                        if (CheckBoxList_Talnakh.Items[i].Selected)
                        {
                            check += CheckBoxList_Talnakh.Items[i].Text;
                        }
                    }
                    if (System.Text.RegularExpressions.Regex.IsMatch(check, "vn", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        if (!check.Contains(target_server))
                        {
                            comboBox.Items.Clear(); //clearing combobox
                            for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                            {
                                if (CheckBoxList_Norilsk.Items[i].Selected)
                                {
                                    //check life server through ping
                                    PingReply reply = pinger.Send(CheckBoxList_Norilsk.Items[i].Text);
                                    pingable = reply.Status == IPStatus.Success;
                                    if (pingable)
                                    {
                                        string script = @"C:\work\Exchange\MaintenanceServer\Start-ExchangeServerMaintenanceMode.ps1 -SourceServer '" + CheckBoxList_Norilsk.Items[i].Text + "' -TargetServerFQDN '" + target_server + ".npr.nornick.ru'";
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
                                                    if (outp.BaseObject.ToString().Contains("INFO: Done! Server " + CheckBoxList_Norilsk.Items[i].Text + " is put succesfully into maintenance mode"))
                                                    {
                                                        this.CheckBoxList_Norilsk.Items[i].Text = CheckBoxList_Norilsk.Items[i].Text.ToUpper();
                                                    }
                                                }
                                            }
                                            ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                            ResultBox.Text += builder.ToString();
                                        }
                                    }
                                    else { ResultBox.Text += "\r\n"+"No ping -> " + CheckBoxList_Norilsk.Items[i].Text; }
                                }
                                comboBox.Items.Add(new ListItem(CheckBoxList_Norilsk.Items[i].Text, CheckBoxList_Norilsk.Items[i].Text));
                            }
                            for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                            {
                                if (CheckBoxList_Talnakh.Items[i].Selected)
                                {
                                    //check life server through ping
                                    PingReply reply = pinger.Send(CheckBoxList_Talnakh.Items[i].Text);
                                    pingable = reply.Status == IPStatus.Success;
                                    if (pingable)
                                    {
                                        script = @"C:\work\Exchange\MaintenanceServer\Start-ExchangeServerMaintenanceMode.ps1 -SourceServer '" + CheckBoxList_Talnakh.Items[i].Text + "' -TargetServerFQDN '" + target_server + ".npr.nornick.ru'";
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
                                                    if (outp.BaseObject.ToString().Contains("INFO: Done! Server " + CheckBoxList_Talnakh.Items[i].Text + " is put succesfully into maintenance mode"))
                                                    {
                                                        this.CheckBoxList_Talnakh.Items[i].Text = CheckBoxList_Talnakh.Items[i].Text.ToUpper();
                                                    }
                                                }
                                            }
                                            ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                            ResultBox.Text += builder.ToString();
                                        }
                                    }
                                    else { ResultBox.Text += "\r\n" + "No ping -> " + CheckBoxList_Talnakh.Items[i].Text; }
                                }
                                comboBox.Items.Add(new ListItem(CheckBoxList_Talnakh.Items[i].Text, CheckBoxList_Talnakh.Items[i].Text));
                            }
                            comboBox.DataBind();
                            File.AppendAllText(path, ResultBox.Text);

                            //update inform to grid-----------------------------------------------------------------------------------------
                            Create_GridTable();
                            Script_Exchange();
                            using (ps = PowerShell.Create())
                            {
                                PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                                ps.AddScript(script);
                                IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                                ps.EndInvoke(result);
                                ps.Stop();
                                check = string.Empty;
                                for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                                {
                                    foreach (PSObject outp in output)
                                    {
                                        get_value_outp(outp);
                                        if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                        {
                                            //check life server through ping
                                            PingReply reply = pinger.Send(NameMailboxServer);
                                            pingable = reply.Status == IPStatus.Success;
                                            check += NameMailboxServer;
                                            Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                        } else if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                        {
                                            //check life server through ping
                                            PingReply reply = pinger.Send(NameMailboxServer);
                                            pingable = reply.Status == IPStatus.Success;
                                            check += NameMailboxServer;
                                            Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber); 
                                        }
                                    }
                                }

                                for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                                {
                                    foreach (PSObject outp in output)
                                    {
                                        get_value_outp(outp);
                                        if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                        {
                                            //check life server through ping
                                            PingReply reply = pinger.Send(NameMailboxServer);
                                            pingable = reply.Status == IPStatus.Success;
                                            check += NameMailboxServer;
                                            Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                        } else if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                        {
                                            //check life server through ping
                                            PingReply reply = pinger.Send(NameMailboxServer);
                                            pingable = reply.Status == IPStatus.Success;
                                            check += NameMailboxServer;
                                            Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber); 
                                        }
                                    }
                                }
                            }
                            GridView1.DataSource = dt;
                            GridView1.DataBind();
                            //update inform to grid-----------------------------------------------------------------------------------------

                            email(ResultBox.Text, "Start Maintenance"); //отправка сообщения
                        }
                        else { ResultBox.Text = "Target Server - " + target_server + " is already exists in Source Server! Choose another Server!"; }
                    }
                    else { ResultBox.Text = "Choose checkbox, please!"; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void Stop_click(object sender, EventArgs e)
        {
            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();

            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
            using (SPLongOperation operation = new SPLongOperation(this.Page))
            {
                string path = @"C:\work\Exchange\MaintenanceServer\Maintenance.log";
                ResultBox.Text = string.Empty;
                if (TextBox1.Text == "Exchange")
                {
                    string check = string.Empty;
                    for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                    {
                        if (CheckBoxList_Norilsk.Items[i].Selected)
                        {
                            check += CheckBoxList_Norilsk.Items[i].Text;
                        }
                    }
                    for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                    {
                        if (CheckBoxList_Talnakh.Items[i].Selected)
                        {
                            check += CheckBoxList_Talnakh.Items[i].Text;
                        }
                    }
                    if (System.Text.RegularExpressions.Regex.IsMatch(check, "vn", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        comboBox.Items.Clear(); //clearing combobox
                        for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                        {
                            if (CheckBoxList_Norilsk.Items[i].Selected)
                            {
                                //check life server through ping
                                PingReply reply = pinger.Send(CheckBoxList_Norilsk.Items[i].Text);
                                pingable = reply.Status == IPStatus.Success;
                                if (pingable)
                                {
                                    script = @"C:\work\Exchange\MaintenanceServer\Stop-ExchangeServerMaintenanceMode.ps1 -Server '" + CheckBoxList_Norilsk.Items[i].Text + "'";
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
                                                if (string.Equals(outp.BaseObject.ToString(), "INFO: Done! Server " + CheckBoxList_Norilsk.Items[i].Text + " successfully taken out of Maintenance Mode.", StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                    this.CheckBoxList_Norilsk.Items[i].Text = CheckBoxList_Norilsk.Items[i].Text.ToLower();
                                                }
                                            }
                                        }
                                        ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                        ResultBox.Text += builder.ToString();
                                    }
                                }
                                else { ResultBox.Text += "\r\n"+"No ping -> " + CheckBoxList_Norilsk.Items[i].Text; }
                            }
                            comboBox.Items.Add(new ListItem(CheckBoxList_Norilsk.Items[i].Text, CheckBoxList_Norilsk.Items[i].Text));
                        }
                        for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                        {
                            if (CheckBoxList_Talnakh.Items[i].Selected)
                            {
                                //check life server through ping
                                PingReply reply = pinger.Send(CheckBoxList_Talnakh.Items[i].Text);
                                pingable = reply.Status == IPStatus.Success;
                                if (pingable)
                                {
                                    script = @"C:\work\Exchange\MaintenanceServer\Stop-ExchangeServerMaintenanceMode.ps1 -Server '" + CheckBoxList_Talnakh.Items[i].Text + "'";
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
                                                if (string.Equals(outp.BaseObject.ToString(), "INFO: Done! Server " + CheckBoxList_Talnakh.Items[i].Text + " successfully taken out of Maintenance Mode.", StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                    this.CheckBoxList_Talnakh.Items[i].Text = CheckBoxList_Talnakh.Items[i].Text.ToLower();
                                                }
                                            }
                                        }
                                        ResultBox.Text += "Username: " + HttpContext.Current.User.Identity.Name.Replace("0#.w|", "") + "\r\n";
                                        ResultBox.Text += builder.ToString();
                                    }
                                }
                                else { ResultBox.Text += "\r\n"+"No ping -> " + CheckBoxList_Talnakh.Items[i].Text; }
                            }
                            comboBox.Items.Add(new ListItem(CheckBoxList_Talnakh.Items[i].Text, CheckBoxList_Talnakh.Items[i].Text));
                        }
                        comboBox.DataBind();
                        File.AppendAllText(path, ResultBox.Text);


                        //update inform to grid-----------------------------------------------------------------------------------------
                        Create_GridTable();
                        Script_Exchange();
                        using (ps = PowerShell.Create())
                        {
                            PSDataCollection<PSObject> output = new PSDataCollection<PSObject>();
                            ps.AddScript(script);
                            IAsyncResult result = ps.BeginInvoke<PSObject, PSObject>(null, output);
                            ps.EndInvoke(result);
                            ps.Stop();
                            check = string.Empty;
                            for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                            {
                                foreach (PSObject outp in output)
                                {
                                    get_value_outp(outp);
                                    if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber); 
                                    }else if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp")){
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber); 
                                    }
                                }
                            }

                            for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                            {
                                foreach (PSObject outp in output)
                                {
                                    get_value_outp(outp);
                                    if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                    {
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber); 
                                    }else if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp")){
                                        //check life server through ping
                                        PingReply reply = pinger.Send(NameMailboxServer);
                                        pingable = reply.Status == IPStatus.Success;
                                        check += NameMailboxServer;
                                        Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber); 
                                    }
                                }
                            }
                        }
                        GridView1.DataSource = dt;
                        GridView1.DataBind();
                        //update inform to grid-----------------------------------------------------------------------------------------
                        email(ResultBox.Text, "Stop Maintenance"); //отправка сообщения
                    }
                    else { ResultBox.Text = "Choose checkbox, please!"; }
                }
                else { ResultBox.Text = "Wrong Input Code!"; }
            }
        }
        protected void Refresh_click(object sender, EventArgs e)
        {
            Create_GridTable();

            //check life server through ping
            bool pingable = false;
            Ping pinger = null;
            pinger = new Ping();
            comboBox.Items.Clear(); //clearing combobox
            Page.Server.ScriptTimeout = 3600; // specify the timeout to 3600 seconds
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
                        //ResultBox.Text = string.Empty;
                        string check = string.Empty;
                        for (int i = 0; i < CheckBoxList_Norilsk.Items.Count; i++)
                        {
                            foreach (PSObject outp in output)
                            {
                                get_value_outp(outp);
                                if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.ToString().Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    this.CheckBoxList_Norilsk.Items[i].Text = CheckBoxList_Norilsk.Items[i].Text.ToUpper();
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                }
                                else if (string.Equals(CheckBoxList_Norilsk.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                }
                            }
                            comboBox.Items.Add(new ListItem(CheckBoxList_Norilsk.Items[i].Text, CheckBoxList_Norilsk.Items[i].Text));
                        }

                        for (int i = 0; i < CheckBoxList_Talnakh.Items.Count; i++)
                        {
                            foreach (PSObject outp in output)
                            {
                                get_value_outp(outp);
                                if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & DatabaseCopyAutoActivationPolicy.Equals("Blocked") & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    this.CheckBoxList_Talnakh.Items[i].Text = CheckBoxList_Talnakh.Items[i].Text.ToUpper();
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToUpper(), "Start", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                }
                                else if (string.Equals(CheckBoxList_Talnakh.Items[i].Text, NameMailboxServer, StringComparison.CurrentCultureIgnoreCase) & !check.Contains(NameMailboxServer) & !NameMailboxServer.Contains("tmp"))
                                {
                                    //check life server through ping
                                    PingReply reply = pinger.Send(NameMailboxServer);
                                    pingable = reply.Status == IPStatus.Success;
                                    check += NameMailboxServer;
                                    Add_to_GridTable(NameMailboxServer.ToLower(), "Stop", pingable.ToString(), DatabaseCopyAutoActivationPolicy, DatabaseCopyActivationDisabledAndMoveNow, MountedDatabase, QueueNumber, MessagesNumber);
                                }
                            }
                            comboBox.Items.Add(new ListItem(CheckBoxList_Talnakh.Items[i].Text, CheckBoxList_Talnakh.Items[i].Text));
                        }
                    }
                    comboBox.DataBind();
                    GridView1.DataSource = dt;
                    GridView1.DataBind();
                }
        }

        void email(string body, string action)
        {
            try
            {
                //File.AppendAllText(@"C:\work\Exchange\MaintenanceServer\Maintenance1.log", body);
                // отправитель - устанавливаем адрес и отображаемое в письме имя
                MailAddress from = new MailAddress("MaintenanceExchange@nornik.ru", "Maintenance Exchange");
                //MailAddress from = new MailAddress("AdmExc@nornik.ru", "AdmExc");
                // кому отправляем
                //MailAddress to = new MailAddress("kordyakim@nornik.ru");
                MailAddress to = new MailAddress("SSC-SCOM-Exchange@nornik.ru");
                // создаем объект сообщения
                MailMessage m = new MailMessage(from, to);
                m.Subject = action;
                m.Body = body;
                // письмо представляет код html
                //m.IsBodyHtml = true;
                SmtpClient smtp = new SmtpClient("smtpzf.npr.nornick.ru", 25);
                //SmtpClient smtp = new SmtpClient("owanr.nornik.ru", 587);
                //smtp.Credentials = new NetworkCredential("AdmExc", "3Xchange");
                //smtp.EnableSsl = true; //do not work
                smtp.Send(m);
            }
            catch (Exception e)
            {}
        }
    }
}
