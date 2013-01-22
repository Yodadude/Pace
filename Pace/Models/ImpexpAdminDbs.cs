using System;
using System.Collections.Generic;
using System.Text;
using NPoco;

namespace Pace.Models
{
	public class DbsRecord
	{
		public string DbType;
		public string Dbms;
		public string Dsn;
		public string UserId;
		public string Pwd;
		public string ServerName;
		public string DbName;
		public string DbParm;
		public string ConnectSettings;
		public string AutoEmailing;
		public string AutoScheduling;
		public int TimeoutLimit;
		public string AutoSchedulerExe;
		public string DefaultSender;
		public string PgpPublicKey;
		public string PgpPassphrase;
		public string PgpPrivateKey;

		public override string ToString()
		{
			return "DSN:" + Dsn + ", Database: " + ServerName + "." + DbName;
		}
	}

	//[TableName("impexp_admin_dbs")]
	//public class DbsRecord
	//{
	//    [Column("dbtype")]
	//    public string Dbtype { get; set; }
	//    [Column("dbms")]
	//    public string Dbms { get; set; }
	//    [Column("dsn")]
	//    public string Dsn { get; set; }
	//    [Column("userid")]
	//    public string Userid { get; set; }
	//    [Column("pwd")]
	//    public string Pwd { get; set; }
	//    [Column("dbname")]
	//    public string Dbname { get; set; }
	//    [Column("dbparm")]
	//    public string Dbparm { get; set; }
	//    [Column("servername")]
	//    public string Servername { get; set; }
	//    [Column("connectsettings")]
	//    public string Connectsettings { get; set; }
	//    [Column("auto_emailing")]
	//    public string AutoEmailing { get; set; }
	//    [Column("auto_scheduling")]
	//    public string AutoScheduling { get; set; }
	//    [Column("timeout_limit")]
	//    public int TimeoutLimit { get; set; }
	//    [Column("auto_scheduler_exe")]
	//    public string AutoSchedulerExe { get; set; }
	//    [Column("default_sender")]
	//    public string DefaultSender { get; set; }
	//    [Column("pgp_public_key")]
	//    public string PgpPublicKey { get; set; }
	//    [Column("pgp_passphrase")]
	//    public string PgpPassphrase { get; set; }
	//    [Column("pgp_private_key")]
	//    public string PgpPrivateKey { get; set; }

	//    public override string ToString()
	//    {
	//        return "DSN:" + Dsn + ", Database: " + Servername + "." + Dbname;
	//    }
	//}
}
