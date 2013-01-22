using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Data.Odbc;
using System.ServiceProcess;
using Pace.Models;

namespace Pace
{
    public partial class Form1 : Form
    {
    	readonly ServiceController _myController;
    	readonly Icon _stopIcon = new Icon(typeof(Form1),"promaster_stop.ico");
    	readonly Icon _startIcon = new Icon(typeof(Form1),"promaster_start.ico");
		Boolean _bRestartService = false;
    	private Boolean _insertMode = false;
    	private DbsRecord _currentDbsRecord = null;


        public Form1()
        {

            InitializeComponent();
            this.toolStripStatusLabel1.Text = "Reading Registry...";

        	EnableParamFields(false);

        	buttonParmSave.Enabled = false;

            String s_odbc_dsn="";
            RegistryKey rkLocalMachine = Registry.LocalMachine;

            // Obtain the key (read-only) and display it.
            RegistryKey rk = rkLocalMachine.OpenSubKey("SOFTWARE\\Inlogik\\Auto_Processor");

            String[] s_names = rk.GetValueNames();

            foreach (String s_keyname in s_names)
            {
                switch (s_keyname.ToLower())
                {
                    case "admin_dbms": break;
                    case "admin_log":
                        textBoxAPLogPath.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "admin_dbtype":
                        this.comboBoxDBType.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "admin_dsn":
                        s_odbc_dsn = rk.GetValue(s_keyname).ToString();
                        break;
                    case "admin_connectsettings":
                        this.textBoxConnectSettings.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "admin_dbparm":
                        this.textBoxDBParm.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "admin_userid":
                        this.textBoxUserID.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "admin_pwd":
                        this.textBoxPassword.Text = rk.GetValue(s_keyname).ToString();
                        break;

                    case "idle_time":
                        this.textBoxIdleTime.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "concurrent_processes":
                        this.textBoxConcurrent.Text = rk.GetValue(s_keyname).ToString();
                        break;

                    case "auto_scheduler_exe":
                        this.textBoxASProgramPath.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "auto_scheduler_log":
                        this.textBoxASLogPath.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "auto_scheduler_flag":
                        this.checkBoxRunAutoSched.Checked = (rk.GetValue(s_keyname).ToString() == "Y");
                        break;

                    case "auto_emailer_exe":
                        this.textBoxAMProgramPath.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "auto_emailer_log":
                        this.textBoxAMLogPath.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "auto_emailer_flag":
                        this.checkBoxRunAutoEmailer.Checked = (rk.GetValue(s_keyname).ToString() == "Y");
                        break;

                    case "email_default_sender":
                        this.textBoxDefaultSender.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "email_smtp":
                        this.textBoxSMTPServer.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "email_port":
                        this.textBoxSMTPPort.Text = rk.GetValue(s_keyname).ToString();
                        break;
                    case "email_html":
                        this.checkBoxHTMLEmail.Checked = (rk.GetValue(s_keyname).ToString() == "Y");
                        break;
                }
            }


            //textBoxAPLogPath.Text = rk.GetValue("Admin_log").ToString();
            //this.comboBoxDBType.Text = rk.GetValue("Admin_dbtype").ToString();
            //s_odbc_dsn = rk.GetValue("Admin_dsn").ToString();
            //this.textBoxConnectSettings.Text = rk.GetValue("Admin_connectsettings").ToString();
            //this.textBoxDBParm.Text = rk.GetValue("Admin_dbparm").ToString();
            //this.textBoxUserID.Text = rk.GetValue("Admin_userid").ToString();
            //this.textBoxPassword.Text = rk.GetValue("Admin_pwd").ToString();
            //this.textBoxIdleTime.Text = rk.GetValue("Idle_time").ToString();

            //this.textBoxASProgramPath.Text = rk.GetValue("Auto_scheduler_exe").ToString();
            //this.textBoxASLogPath.Text = rk.GetValue("Auto_scheduler_log").ToString();
            //this.checkBoxRunAutoSched.Checked = (rk.GetValue("Auto_scheduler_flag").ToString() == "Y");

            //this.textBoxAMProgramPath.Text = rk.GetValue("Auto_emailer_exe").ToString();
            //this.textBoxAMLogPath.Text = rk.GetValue("Auto_emailer_log").ToString();
            //this.checkBoxRunAutoEmailer.Checked = (rk.GetValue("Auto_emailer_flag").ToString() == "Y");

            //if (rk.SubKeyCount > 14)
            //{
            //    this.textBoxDefaultSender.Text = rk.GetValue("email_default_sender").ToString();
            //    this.textBoxSMTPServer.Text = rk.GetValue("email_smtp").ToString();
            //    this.textBoxSMTPPort.Text = rk.GetValue("email_port").ToString();
            //    this.checkBoxHTMLEmail.Checked = (rk.GetValue("email_html").ToString() == "Y");
            //}

        	PopulateOdbcDsn(ref comboBoxDSN);

            comboBoxDSN.SelectedItem = s_odbc_dsn;

            this.toolStripStatusLabel1.Text = "Ready";

            String status;

            _myController = new ServiceController("Promaster Auto Processor");

            status = _myController.Status.ToString();
            labelServiceStatus.Text = status;

			if (status.Equals("Running"))
			{
				buttonStop.Enabled = true;
				toolStripMenuItemStop.Enabled = true;
				toolStripMenuItemStart.Enabled = false;
				toolStripMenuItemRestart.Enabled = true;
				notifyIcon1.Icon = _startIcon;
			}
			else if (status.Equals("Stopped"))
			{
				buttonStart.Enabled = true;
				toolStripMenuItemStart.Enabled = true;
				toolStripMenuItemStop.Enabled = false;
				toolStripMenuItemRestart.Enabled = false;
				notifyIcon1.Icon = _stopIcon;
			}
	
			this.Visible = false;

        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonTestConnection_Click(object sender, EventArgs e)
        {
            String connectionString;
            connectionString = "DSN="+ comboBoxDSN.SelectedItem.ToString()+";uid="+textBoxUserID.Text+";pwd="+textBoxPassword.Text;
            this.toolStripStatusLabel1.Text = "Testing connection...";
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    connection.Open();
                }
                MessageBox.Show("Connection successful", "Test connection");
            }
            catch (System.Data.Odbc.OdbcException err)
            {
                MessageBox.Show(err.Message, "Unable to connect");
            }
            this.toolStripStatusLabel1.Text = "Ready";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = textBoxAPLogPath.Text.Substring(1, textBoxAPLogPath.Text.LastIndexOf('\\'));
            openFileDialog1.FileName = textBoxAPLogPath.Text;
            openFileDialog1.Filter = "log files (*.log)|*.log|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxAPLogPath.Text = openFileDialog1.FileName;
            }

        }

        private void buttonASProgPath_Click(object sender, EventArgs e)
        {

            if (textBoxASProgramPath.Text.Length > 0)
            {
                openFileDialog1.InitialDirectory = textBoxASProgramPath.Text.Substring(1, textBoxASProgramPath.Text.LastIndexOf('\\'));
                openFileDialog1.FileName = textBoxASProgramPath.Text;
            }
            openFileDialog1.Filter = "exe files (*.exe)|*.exe|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxASProgramPath.Text = openFileDialog1.FileName;
            }

        }

        private void buttonASLogPath_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = textBoxASLogPath.Text.Substring(1, textBoxASLogPath.Text.LastIndexOf('\\'));
            openFileDialog1.FileName = textBoxASLogPath.Text;
            openFileDialog1.Filter = "log files (*.log)|*.log|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxASLogPath.Text = openFileDialog1.FileName;
            }
        }

        private void buttonAMProgPath_Click(object sender, EventArgs e)
        {
            if (textBoxAMProgramPath.Text.Length > 0)
            {
                openFileDialog1.InitialDirectory = textBoxAMProgramPath.Text.Substring(1, textBoxAMProgramPath.Text.LastIndexOf('\\'));
                openFileDialog1.FileName = textBoxAMProgramPath.Text;
            }
            openFileDialog1.Filter = "exe files (*.exe)|*.exe|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxAMProgramPath.Text = openFileDialog1.FileName;
            }
        }

        private void buttonAMLogPath_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = textBoxAMLogPath.Text.Substring(1, textBoxAMLogPath.Text.LastIndexOf('\\'));
            openFileDialog1.FileName = textBoxAMLogPath.Text;
            openFileDialog1.Filter = "log files (*.log)|*.log|All files (*.*)|*.*";
            openFileDialog1.CheckFileExists = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxAMLogPath.Text = openFileDialog1.FileName;
            }
        }

        private void buttonSaveRegistry_Click(object sender, EventArgs e)
        {
            RegistryKey rkLocalMachine = Registry.LocalMachine;
            RegistryKey rk = rkLocalMachine.OpenSubKey("SOFTWARE\\Inlogik\\Auto_Processor", true);
            String s_auto_scheduler_flag="N";
            String s_auto_emailer_flag="N", s_email_html="N";

            if (checkBoxRunAutoSched.Checked) s_auto_scheduler_flag="Y";
            if (checkBoxRunAutoEmailer.Checked) s_auto_emailer_flag="Y";
            if (checkBoxHTMLEmail.Checked) s_email_html = "Y";

            try
            {
                String[] s_names = rk.GetValueNames();

                foreach (String s_keyname in s_names)
                {
                    switch (s_keyname.ToLower())
                    {
                        case "admin_dbms": break;
                        case "admin_log":
                            rk.SetValue(s_keyname, textBoxAPLogPath.Text.ToString());
                            break;
                        case "admin_dbtype":
                            rk.SetValue(s_keyname, comboBoxDBType.Text.ToString());
                            break;
                        case "admin_dsn":
                            rk.SetValue(s_keyname, comboBoxDSN.SelectedItem.ToString());
                            break;
                        case "admin_connectsettings":
                            rk.SetValue(s_keyname, textBoxConnectSettings.Text.ToString());
                            break;
                        case "admin_dbparm":
                            rk.SetValue(s_keyname, textBoxDBParm.Text.ToString());
                            break;
                        case "admin_userid":
                            rk.SetValue(s_keyname, textBoxUserID.Text.ToString());
                            break;
                        case "admin_pwd":
                            rk.SetValue(s_keyname, textBoxPassword.Text.ToString());
                            break;

                        case "idle_time":
                            rk.SetValue(s_keyname, textBoxIdleTime.Text.ToString());
                            break;
                        case "concurrent_processes":
                            rk.SetValue(s_keyname, this.textBoxConcurrent.Text.ToString());
                            break;

                        case "auto_scheduler_exe":
                            rk.SetValue(s_keyname, textBoxASProgramPath.Text.ToString());
                            break;
                        case "auto_scheduler_log":
                            rk.SetValue(s_keyname, textBoxASLogPath.Text.ToString());
                            break;
                        case "auto_scheduler_flag":
                            rk.SetValue(s_keyname, s_auto_scheduler_flag);
                            break;

                        case "auto_emailer_exe":
                            rk.SetValue(s_keyname, textBoxAMProgramPath.Text.ToString());
                            break;
                        case "auto_emailer_log":
                            rk.SetValue(s_keyname, textBoxAMLogPath.Text);
                            break;
                        case "auto_emailer_flag":
                            rk.SetValue(s_keyname, s_auto_emailer_flag);
                            break;

                        case "email_default_sender":
                            rk.SetValue(s_keyname, textBoxDefaultSender.Text.ToString());
                            break;
                        case "email_smtp":
                            rk.SetValue(s_keyname, textBoxSMTPServer.Text.ToString());
                            break;
                        case "email_port":
                            rk.SetValue(s_keyname, textBoxSMTPPort.Text.ToString());
                            break;
                        case "email_html":
                            rk.SetValue(s_keyname, s_email_html);
                            break;
                    }
                }

            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message, "Error writting to Registry");
            }
            finally
            {
                rk.Close();
                rkLocalMachine.Close();
                this.toolStripStatusLabel1.Text = "Auto Processor settings successfully saved.";
            }

        }

        private void buttonRetrieve_Click(object sender, EventArgs e)
        {
            int i;
            String connectionString;
            connectionString = "DSN=" + comboBoxDSN.SelectedItem.ToString() + ";uid=" + textBoxUserID.Text + ";pwd=" + textBoxPassword.Text;

            this.dataGridView1.Rows.Clear();

            try
            {
                OdbcConnection conn = new OdbcConnection(connectionString);
                OdbcCommand cmd = new OdbcCommand();
                cmd.CommandText = "select dbtype,dbms,dsn,userid,pwd,dbparm,connectsettings,auto_emailing,auto_scheduling,timeout_limit,auto_scheduler_exe,default_sender from impexp_admin_dbs";
                cmd.Connection = conn;
                conn.Open();
                OdbcDataReader data = cmd.ExecuteReader();

                while (data.Read())
                {

                    i = this.dataGridView1.Rows.Add();
                    this.dataGridView1["dbtype", i].Value = data.GetString(0);
                    this.dataGridView1["dbms", i].Value = data.GetString(1);
                    this.dataGridView1["dsn", i].Value = data.GetString(2);
                    this.dataGridView1["userid", i].Value = data.GetString(3);
                    this.dataGridView1["password", i].Value = data.GetString(4);
                    if (!data.IsDBNull(5))
                        this.dataGridView1["dbparm", i].Value = data.GetString(5);

                    this.dataGridView1["connectsettings", i].Value = data.GetString(6);
                    this.dataGridView1["auto_emailing", i].Value = data.GetString(7);
                    this.dataGridView1["auto_scheduling", i].Value = data.GetString(8);
                    this.dataGridView1["timeout_limit", i].Value = data.GetInt32(9);
                    if (!data.IsDBNull(10))
                        this.dataGridView1["auto_scheduler_exe", i].Value = data.GetString(10);
                    if (!data.IsDBNull(11))
                        this.dataGridView1["default_sender", i].Value = data.GetString(11);
                }
                data.Close();
                conn.Close();
            }
            catch (SystemException err)
            {
                MessageBox.Show(err.Message, this.Text);
            }
        }

        private Boolean validateConnectionData()
        {
            // validate data for all rows
            int i;
            Boolean b_errors = false;

            for (i = 0; i < this.dataGridView1.RowCount; i++)
            {
                if (this.dataGridView1.Rows[i].Cells["dsn"].Value == null)
                {
                    MessageBox.Show("ODBC DSN not specified.", this.Text);
                    b_errors = true;
                    break;
                }
                if (this.dataGridView1.Rows[i].Cells["userid"].Value == null)
                {
                    MessageBox.Show("User ID is not specified.", this.Text);
                    b_errors = true;
                    break;
                }
                if (this.dataGridView1.Rows[i].Cells["password"].Value == null)
                {
                    MessageBox.Show("Password not specified.", this.Text);
                    b_errors = true;
                    break;
                }
                if (this.dataGridView1.Rows[i].Cells["timeout_limit"].Value == null)
                {
                    MessageBox.Show("Timeout Limit not specified.", this.Text);
                    b_errors = true;
                    break;
                }
            }

            return b_errors;

        }

        private void buttonSaveDB_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.RowCount == 0) return;
            if (validateConnectionData()) return;

            int i;
            String connectionString;
            connectionString = "DSN=" + comboBoxDSN.SelectedItem.ToString() + ";uid=" + textBoxUserID.Text + ";pwd=" + textBoxPassword.Text;

            OdbcConnection conn = new OdbcConnection(connectionString);
            OdbcTransaction transaction = null;
            OdbcCommand cmd = new OdbcCommand();
            cmd.Connection = conn;

            try
            {
                conn.Open();
                transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;
                cmd.CommandText = "delete from impexp_admin_dbs";
                cmd.ExecuteNonQuery();
                cmd.Transaction = transaction;
                cmd.Parameters.Clear();
                cmd.CommandText = "insert into impexp_admin_dbs " +
                                  "(dbtype,dbms,dsn,userid,pwd,dbparm,connectsettings,auto_emailing,auto_scheduling,timeout_limit,auto_scheduler_exe,default_sender)" +
                                  " values(?,?,?,?,?,?,?,?,?,?,?,?)";
                cmd.Parameters.Add("@dbtype", OdbcType.VarChar, 20);
                cmd.Parameters.Add("@dbms", OdbcType.VarChar, 20);
                cmd.Parameters.Add("@dsn", OdbcType.VarChar, 100);
                cmd.Parameters.Add("@userid", OdbcType.VarChar, 40);
                cmd.Parameters.Add("@pwd", OdbcType.VarChar, 40);
                cmd.Parameters.Add("@dbparm", OdbcType.VarChar, 255);
                cmd.Parameters.Add("@csett", OdbcType.VarChar, 255);
                cmd.Parameters.Add("@emailer", OdbcType.VarChar, 1);
                cmd.Parameters.Add("@scheduler", OdbcType.VarChar, 1);
                cmd.Parameters.Add("@timeout", OdbcType.Int, 0);
                cmd.Parameters.Add("@exe_path", OdbcType.VarChar, 255);
                cmd.Parameters.Add("@default_sender", OdbcType.VarChar, 255);


                for (i = 0; i < this.dataGridView1.RowCount; i++)
                {
                    cmd.Parameters["@dbtype"].Value = this.dataGridView1["dbtype", i].Value.ToString();
                    cmd.Parameters["@dbms"].Value = this.dataGridView1["dbms", i].Value.ToString();
                    cmd.Parameters["@dsn"].Value = this.dataGridView1["dsn", i].Value.ToString();
                    cmd.Parameters["@userid"].Value = this.dataGridView1["userid", i].Value.ToString();
                    cmd.Parameters["@pwd"].Value = this.dataGridView1["password", i].Value.ToString();

                    if (this.dataGridView1["dbparm", i].Value == null)
                        cmd.Parameters["@dbparm"].Value = Convert.DBNull;
                    else
                        cmd.Parameters["@dbparm"].Value = this.dataGridView1["dbparm", i].Value.ToString();

                    cmd.Parameters["@csett"].Value = this.dataGridView1["connectsettings", i].Value.ToString();
                    cmd.Parameters["@emailer"].Value = this.dataGridView1["auto_emailing", i].Value;
                    cmd.Parameters["@scheduler"].Value = this.dataGridView1["auto_scheduling", i].Value;
                    cmd.Parameters["@timeout"].Value = this.dataGridView1["timeout_limit", i].Value;

                    if (this.dataGridView1["auto_scheduler_exe", i].Value == null)
                        cmd.Parameters["@exe_path"].Value = Convert.DBNull;
                    else
                        cmd.Parameters["@exe_path"].Value = this.dataGridView1["auto_scheduler_exe", i].Value.ToString();

                    if (this.dataGridView1["default_sender", i].Value == null)
                        cmd.Parameters["@default_sender"].Value = Convert.DBNull;
                    else
                        cmd.Parameters["@default_sender"].Value = this.dataGridView1["default_sender", i].Value;

                    cmd.ExecuteNonQuery();
                }

                transaction.Commit();

                this.toolStripStatusLabel1.Text = "Connection settings successfully saved.";
            }
            catch (System.Exception e1)
            {
                MessageBox.Show(e1.Message, this.Text);
                try
                {
                    transaction.Rollback();
                }
                catch (System.Exception e2)
                {
                    MessageBox.Show(e2.Message, this.Text);
                }
            }
            finally
            {
                cmd.Dispose();
                transaction.Dispose();
                conn.Close();
            }

        }

        private void buttonAddRow_Click(object sender, EventArgs e)
        {
            int i;
            i = this.dataGridView1.Rows.Add();
            this.dataGridView1["dbtype", i].Value = "ORACLE";
            this.dataGridView1["dbms", i].Value = "ODBC";
            this.dataGridView1["userid", i].Value = "promaster";
            this.dataGridView1["password", i].Value = "changeoninstall";
            this.dataGridView1["connectsettings", i].Value = ",DelimitIdentifier='No',DisableBind=1,StaticBind=0,CommitOnDisconnect='No',ConnectOption='SQL_DRIVER_CONNECT,SQL_DRIVER_NOPROMPT',CommitOnDisconnect='No'";
            this.dataGridView1["auto_emailing", i].Value = "Y";
            this.dataGridView1["auto_scheduling", i].Value = "Y";
            this.dataGridView1["timeout_limit", i].Value = "10";
        }

        private void buttonDeleteRow_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.RowCount > 0)
            {
                this.dataGridView1.Rows.RemoveAt(this.dataGridView1.CurrentRow.Index);
            }
        }

        private void buttonTestAllConnections_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.RowCount == 0) return;
            if (validateConnectionData()) return;

            int i;
            String connectionString, dsn="", userid="", pwd="";

            OdbcConnection conn = new OdbcConnection();

            try
            {
				//for (i = 0; i < this.dataGridView1.RowCount; i++)
				//{
				//    dsn = this.dataGridView1["dsn", i].Value.ToString();
				//    userid = this.dataGridView1["userid", i].Value.ToString();
				//    pwd = this.dataGridView1["password", i].Value.ToString();
				//    connectionString = "DSN=" + dsn + ";uid=" + userid + ";pwd=" + pwd;
				//    conn.ConnectionString = connectionString;
				//    conn.Open();
				//    conn.Close();
				//}
				foreach (DbsRecord item in comboBoxDBList.Items)
            	{
					connectionString = "DSN=" + item.Dsn + ";uid=" + item.UserId + ";pwd=" + item.Pwd;
					conn.ConnectionString = connectionString;
					conn.Open();
					conn.Close();
            	}
            }
            catch (System.Exception e1)
            {
                MessageBox.Show(e1.Message+"\r\n\r\nDSN="+dsn+"\r\nUserID="+userid+"\r\nPassword="+pwd, this.Text);
            }
            finally
            {
                conn.Close();
                this.toolStripStatusLabel1.Text = "All connections successfully connected.";
            }
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
			startAutoScheduler();
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
			stopAutoScheduler();
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            _myController.Refresh();
            labelServiceStatus.Text = _myController.Status.ToString();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //MessageBox.Show(e.Exception.Message);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            _myController.Refresh();

            if (_myController.Status.ToString().Equals("Running"))
            {
                timer1.Stop();
				hasStarted();
                
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            _myController.Refresh();

            if (_myController.Status.ToString().Equals("Stopped"))
            {
                timer2.Stop();
				hasStopped();

				if (_bRestartService)
				{
					startAutoScheduler();
				}
            }
        }

		private void hasStarted()
		{
			labelServiceStatus.Text = _myController.Status.ToString();
			pictureBox1.Visible = false;
			buttonStart.Enabled = false;
			buttonStop.Enabled = true;

			toolStripMenuItemStop.Enabled = true;
			toolStripMenuItemStart.Enabled = false;
			toolStripMenuItemRestart.Enabled = true;
			notifyIcon1.Icon = _startIcon;
			_bRestartService = false;
		}

		private void hasStopped()
		{
			labelServiceStatus.Text = _myController.Status.ToString();
			pictureBox1.Visible = false;
			buttonStart.Enabled = true;
			buttonStop.Enabled = false;

			toolStripMenuItemStart.Enabled = true;
			toolStripMenuItemStop.Enabled = false;
			toolStripMenuItemRestart.Enabled = false;
			notifyIcon1.Icon = _stopIcon;
		}

		private void Form1_Resize(object sender, EventArgs e)
		{
			if (this.WindowState == FormWindowState.Minimized)
			{
				this.Visible = false;
				this.notifyIcon1.Visible = true;
			}
		}

		private void Display()
		{
			this.Visible = true;
            //this.notifyIcon1.Visible = false;
			this.WindowState = FormWindowState.Normal;
			this.Activate();
		}

		private void notifyIcon1_DoubleClick(object sender, EventArgs e)
		{
			this.Display();
		}

		private void toolStripMenuItemExit_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}

		private void toolStripMenuItemEdit_Click(object sender, EventArgs e)
		{
			this.Display();
		}


		private void startAutoScheduler()
		{
			try
			{
				_myController.Start();
				_myController.Refresh();
				labelServiceStatus.Text = _myController.Status.ToString();

				if (!_myController.Status.ToString().Equals("Running"))
				{
					pictureBox1.Visible = true;
					timer1.Start();
				}
				else
				{
					hasStarted();
				}

			}
			catch (System.Exception e1)
			{
				MessageBox.Show(e1.Message);
			}
		}

		private void stopAutoScheduler()
		{
			try
			{
				if (_myController.CanStop)
				{
					_myController.Stop();
					_myController.Refresh();
					labelServiceStatus.Text = _myController.Status.ToString();
					if (!_myController.Status.ToString().Equals("Stopped"))
					{
						pictureBox1.Visible = true;
						timer2.Start();
					}
					else
					{
						hasStopped();
					}
				}
			}
			catch (System.Exception e1)
			{
				MessageBox.Show(e1.Message);
			}
		}

		private void toolStripMenuItemStart_Click(object sender, EventArgs e)
		{
			startAutoScheduler();
		}

		private void toolStripMenuItemStop_Click(object sender, EventArgs e)
		{
			stopAutoScheduler();
		}

		private void toolStripMenuItemRestart_Click(object sender, EventArgs e)
		{
			_bRestartService = true;
			stopAutoScheduler();
		}

		private void tabPageDatabases_Enter(object sender, EventArgs e)
		{

			if (comboBoxDBList.Items.Count > 0)
				return;

			string connectionString = string.Format("DSN={0};uid={1};pwd={2};", comboBoxDSN.SelectedItem, textBoxUserID.Text, textBoxPassword.Text);
			
			try
			{
				var conn = new OdbcConnection(connectionString);
				var cmd = new OdbcCommand
				          	{
				          		CommandText = @"select dbtype,dbms,dsn,userid,pwd,servername,dbname,dbparm,connectsettings,auto_emailing,auto_scheduling,auto_scheduler_exe,timeout_limit,default_sender,pgp_public_key,pgp_passphrase,pgp_private_key from impexp_admin_dbs",
				          		Connection = conn
				          	};
				conn.Open();
				var data = cmd.ExecuteReader();

				while (data.Read())
				{
					comboBoxDBList.Items.Add(
						new DbsRecord
							{
								DbType = data.IsDBNull(0) ? "SQL SERVER" : data.GetString(0),
								Dbms = data.IsDBNull(1) ? "ODBC" : data.GetString(1),
								Dsn = data.IsDBNull(2) ? "" : data.GetString(2),
								UserId = data.IsDBNull(3) ? "promaster" : data.GetString(3),
								Pwd = data.IsDBNull(4) ? "" : data.GetString(4),
								ServerName = data.IsDBNull(5) ? "" : data.GetString(5),
								DbName = data.IsDBNull(6) ? "" : data.GetString(6),
								DbParm = data.IsDBNull(7) ? "" : data.GetString(7),
								ConnectSettings = data.IsDBNull(8) ? ",DelimitIdentifier='No',DisableBind=1,StaticBind=0,ConnectOption='SQL_DRIVER_CONNECT,SQL_DRIVER_NOPROMPT',CommitOnDisconnect='No'" : data.GetString(8),
								AutoEmailing = data.IsDBNull(9) ? "N" : data.GetString(9),
								AutoScheduling = data.IsDBNull(10) ? "N" : data.GetString(10),
								AutoSchedulerExe = data.IsDBNull(11) ? "" : data.GetString(11),
								TimeoutLimit = data.IsDBNull(12) ? 0:data.GetInt32(12),
								DefaultSender = data.IsDBNull(13) ? "" :data.GetString(13),
								PgpPublicKey = data.IsDBNull(14) ? "": data.GetString(14),
								PgpPassphrase = data.IsDBNull(15) ? "":data.GetString(15),
								PgpPrivateKey = data.IsDBNull(16) ? "":data.GetString(16)
							}
						);

				}
				data.Close();
				conn.Close();
			}
			catch (SystemException err)
			{
				MessageBox.Show(err.Message, this.Text);
			}
		}

		private void comboBoxDBList_SelectedIndexChanged(object sender, EventArgs e)
		{
			var comboBox = (ComboBox)sender;
			_currentDbsRecord = (DbsRecord)comboBox.SelectedItem;

			PopulateOdbcDsn(ref comboBoxParmDSN);

			comboBoxParmDBType.Text = _currentDbsRecord.DbType;
			comboBoxParmDBMS.Text = _currentDbsRecord.Dbms;
			comboBoxParmDSN.Text = _currentDbsRecord.Dsn;
			textBoxParmUserId.Text = _currentDbsRecord.UserId;
			textBoxParmPassword.Text = _currentDbsRecord.Pwd;
			textBoxParmServerName.Text = _currentDbsRecord.ServerName;
			textBoxParmDatabase.Text = _currentDbsRecord.DbName;
			textBoxParmDBParm.Text = _currentDbsRecord.DbParm;
			textBoxParmConnectSettings.Text = _currentDbsRecord.ConnectSettings;
			checkBoxParmAutoScheduling.Checked = _currentDbsRecord.AutoScheduling == "Y";
			checkBoxParmAutoEmailing.Checked = _currentDbsRecord.AutoEmailing == "Y";
			textBoxParmDefaultSender.Text = _currentDbsRecord.DefaultSender;
			textBoxParmTimeout.Text = _currentDbsRecord.TimeoutLimit.ToString();
			textBoxParmPGPPublicKey.Text = _currentDbsRecord.PgpPublicKey;
			textBoxParmPGPPassphrase.Text = _currentDbsRecord.PgpPassphrase;
			textBoxParmPGPPrivateKey.Text = _currentDbsRecord.PgpPrivateKey;

			_insertMode = false;

			EnableParamFields(true);
		}


		private void PopulateOdbcDsn(ref ComboBox combobox)
		{
			combobox.Items.Clear();

			RegistryKey rkLocalMachine = Registry.LocalMachine;
			RegistryKey rk = rkLocalMachine.OpenSubKey("Software\\ODBC\\ODBC.INI\\ODBC Data Sources");

			foreach (string subKeyName in rk.GetValueNames())
			{
				combobox.Items.Add(subKeyName);
			}

			rk.Close();
			rkLocalMachine.Close();

		}

		private void buttonParmSave_Click(object sender, EventArgs e)
		{
			int timeLimit;
			if (!int.TryParse(textBoxParmTimeout.Text, out timeLimit))
			{
				MessageBox.Show("Timeout Limit is invalid.");
				textBoxParmTimeout.Focus();
				return;
			}

			var editedDbsRecord = new DbsRecord
			{
				DbType = comboBoxParmDBType.SelectedItem.ToString(),
				Dbms = comboBoxParmDBMS.SelectedItem.ToString(),
				Dsn = comboBoxParmDSN.SelectedItem.ToString(),
				UserId = textBoxParmUserId.Text,
				Pwd = textBoxParmPassword.Text,
				ServerName = textBoxParmServerName.Text,
				DbName = textBoxParmDatabase.Text,
				DbParm = textBoxParmDBParm.Text,
				ConnectSettings = textBoxParmConnectSettings.Text,
				AutoScheduling = checkBoxParmAutoScheduling.Checked ? "Y" : "N",
				AutoEmailing = checkBoxParmAutoEmailing.Checked ? "Y" : "N",
				AutoSchedulerExe = textBoxParmAutoSchedulerPath.Text,
				DefaultSender = textBoxParmDefaultSender.Text,
				TimeoutLimit = Convert.ToInt32(textBoxParmTimeout.Text),
				PgpPublicKey = textBoxParmPGPPublicKey.Text,
				PgpPassphrase = textBoxParmPGPPassphrase.Text,
				PgpPrivateKey = textBoxParmPGPPrivateKey.Text
			};

			if (_insertMode)
			{
				comboBoxDBList.SelectedIndex = comboBoxDBList.Items.Add(editedDbsRecord);

			}
			else
			{
				comboBoxDBList.Items[comboBoxDBList.SelectedIndex] = editedDbsRecord;
			}

			SaveParameters();

			_insertMode = false;
		}

		private void SaveParameters()
		{
			string connectionString = GetConnectionString();

			var conn = new OdbcConnection(connectionString);
			OdbcTransaction transaction = null;
			var cmd = new OdbcCommand();
			cmd.Connection = conn;

			try
			{
				conn.Open();
				transaction = conn.BeginTransaction();
				cmd.Transaction = transaction;

				cmd.CommandText = "delete from impexp_admin_dbs";

				cmd.ExecuteNonQuery();

				cmd.Parameters.Clear();

				cmd.Transaction = transaction;

				cmd.CommandText = "insert into impexp_admin_dbs " +
								  "(dbtype,dbms,dsn,userid,pwd,servername,dbname,dbparm,connectsettings,auto_emailing,auto_scheduling,timeout_limit,auto_scheduler_exe,default_sender,pgp_public_key,pgp_passphrase,pgp_private_key)" +
								  " values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

				cmd.Parameters.Add("@dbtype", OdbcType.VarChar, 20);
				cmd.Parameters.Add("@dbms", OdbcType.VarChar, 20);
				cmd.Parameters.Add("@dsn", OdbcType.VarChar, 100);
				cmd.Parameters.Add("@userid", OdbcType.VarChar, 40);
				cmd.Parameters.Add("@pwd", OdbcType.VarChar, 40);
				cmd.Parameters.Add("@servername", OdbcType.VarChar, 100);
				cmd.Parameters.Add("@dbname", OdbcType.VarChar, 100);
				cmd.Parameters.Add("@dbparm", OdbcType.VarChar, 255);
				cmd.Parameters.Add("@consett", OdbcType.VarChar, 255);
				cmd.Parameters.Add("@emailer", OdbcType.VarChar, 1);
				cmd.Parameters.Add("@scheduler", OdbcType.VarChar, 1);
				cmd.Parameters.Add("@timeout", OdbcType.Int, 0);
				cmd.Parameters.Add("@exe_path", OdbcType.VarChar, 255);
				cmd.Parameters.Add("@default_sender", OdbcType.VarChar, 255);
				cmd.Parameters.Add("@pgp_public_key", OdbcType.VarChar, 3000);
				cmd.Parameters.Add("@pgp_passphrase", OdbcType.VarChar, 100);
				cmd.Parameters.Add("@pgp_private_key", OdbcType.VarChar, 2060);

				foreach (DbsRecord item in comboBoxDBList.Items)
				{
					cmd.Parameters["@dbtype"].Value = item.DbType;
					cmd.Parameters["@dbms"].Value = item.Dbms;
					cmd.Parameters["@dsn"].Value = item.Dsn;
					cmd.Parameters["@userid"].Value = item.UserId.Left(40);
					cmd.Parameters["@pwd"].Value = item.Pwd.Left(40);
					cmd.Parameters["@servername"].Value = item.ServerName.Left(100);
					cmd.Parameters["@dbname"].Value = item.DbName.Left(100);
					cmd.Parameters["@dbparm"].Value = item.DbParm.Left(255);
					cmd.Parameters["@consett"].Value = item.ConnectSettings != null ? item.ConnectSettings.Left(255) : "";
					cmd.Parameters["@emailer"].Value = item.AutoEmailing.Left(1);
					cmd.Parameters["@scheduler"].Value = item.AutoScheduling.Left(1);
					cmd.Parameters["@timeout"].Value = item.TimeoutLimit;
					cmd.Parameters["@exe_path"].Value = item.AutoSchedulerExe.Left(255);
					cmd.Parameters["@default_sender"].Value = item.DefaultSender != null ? item.DefaultSender.Left(255) : "";
					cmd.Parameters["@pgp_public_key"].Value = item.PgpPublicKey.Left(3000);
					cmd.Parameters["@pgp_passphrase"].Value = item.PgpPassphrase.Left(100);
					cmd.Parameters["@pgp_private_key"].Value = item.PgpPrivateKey.Left(2060);

					cmd.ExecuteNonQuery();
				}

				transaction.Commit();

				this.toolStripStatusLabel1.Text = "Connection settings successfully saved.";
			}
			catch (Exception e1)
			{
				MessageBox.Show(e1.Message, this.Text);
				try
				{
					transaction.Rollback();
				}
				catch (Exception e2)
				{
					MessageBox.Show(e2.Message, this.Text);
				}
			}
			finally
			{
				cmd.Dispose();
				transaction.Dispose();
				conn.Close();
			}

		}

		private string GetConnectionString()
		{
			return "DSN=" + comboBoxDSN.SelectedItem + ";uid=" + textBoxUserID.Text + ";pwd=" + textBoxPassword.Text;
		}

		private void buttonNewParam_Click(object sender, EventArgs e)
		{
			
			_insertMode = true;
			ResetParamFields();
			EnableParamFields(true);
		}

		private void ResetParamFields()
		{
			PopulateOdbcDsn(ref comboBoxParmDSN);

			comboBoxParmDBType.SelectedIndex = 0;
			comboBoxParmDBMS.SelectedIndex = 0;
			comboBoxParmDSN.SelectedIndex = 0;
			textBoxParmUserId.Text = "";
			textBoxParmPassword.Text = "";
			textBoxParmServerName.Text = "";
			textBoxParmDatabase.Text = "";
			textBoxParmDBParm.Text = "";
			textBoxParmConnectSettings.Text = ",DelimitIdentifier='No',DisableBind=1,StaticBind=0,ConnectOption='SQL_DRIVER_CONNECT,SQL_DRIVER_NOPROMPT',CommitOnDisconnect='No'";
			checkBoxParmAutoScheduling.Checked = true;
			checkBoxParmAutoEmailing.Checked = true;
			textBoxParmDefaultSender.Text = "support@inlogik.com";
			textBoxParmTimeout.Text = "10";
			textBoxParmPGPPublicKey.Text = "";
			textBoxParmPGPPassphrase.Text = "";
			textBoxParmPGPPrivateKey.Text = "";
		}

		private void EnableParamFields(Boolean protect)
		{
			comboBoxParmDBType.Enabled = protect;
			comboBoxParmDBMS.Enabled = protect;
			comboBoxParmDSN.Enabled = protect;
			textBoxParmUserId.Enabled = protect;
			textBoxParmPassword.Enabled = protect;
			textBoxParmServerName.Enabled = protect;
			textBoxParmDatabase.Enabled = protect;
			textBoxParmDBParm.Enabled = protect;
			textBoxParmConnectSettings.Enabled = protect;
			checkBoxParmAutoScheduling.Enabled = protect;
			textBoxParmAutoSchedulerPath.Enabled = protect;
			checkBoxParmAutoEmailing.Enabled = protect;
			textBoxParmDefaultSender.Enabled = protect;
			textBoxParmTimeout.Enabled = protect;
			textBoxParmPGPPublicKey.Enabled = protect;
			textBoxParmPGPPassphrase.Enabled = protect;
			textBoxParmPGPPrivateKey.Enabled = protect;
			buttonParmSave.Enabled = protect;
		}

		private void buttonParamDelete_Click(object sender, EventArgs e)
		{
			if (comboBoxDBList.SelectedIndex >= 0)
			{
				comboBoxDBList.Items.RemoveAt(comboBoxDBList.SelectedIndex);
				SaveParameters();
				if (comboBoxDBList.Items.Count == 0)
				{
					ResetParamFields();
					EnableParamFields(false);
				}
			}
		}

		private void buttonTestAll_Click(object sender, EventArgs e)
		{

			var connectionString= "";
			var conn = new OdbcConnection();

			try
			{
				foreach (DbsRecord item in comboBoxDBList.Items)
				{
					connectionString = "DSN=" + item.Dsn + ";uid=" + item.UserId + ";pwd=" + item.Pwd;
					conn.ConnectionString = connectionString;
					conn.Open();
					conn.Close();
				}
			}
			catch (Exception e1)
			{
				MessageBox.Show(e1.Message + "\r\n\r\n" + connectionString, this.Text);
			}
			finally
			{
				conn.Close();
				this.toolStripStatusLabel1.Text = "All connections successfully connected.";
			}
		}

    }

}