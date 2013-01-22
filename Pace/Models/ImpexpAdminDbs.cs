
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

}
