using System;
using System.Configuration;
using Backup_Db_ATD.Repositorys;

namespace Backup_Db_ATD
{
    class Program
    {
        static void Main(string[] args)
        {
            Repo_SQL repo_SQL = new Repo_SQL();
            try
            {
                string v_loai_csdl = ConfigurationManager.AppSettings["LoaiCSDL"].ToString();
                string v_server = ConfigurationManager.AppSettings["Server"].ToString();
                string v_ten_csdl = ConfigurationManager.AppSettings["TenCSDL"].ToString();
                string v_loai_saoluu = ConfigurationManager.AppSettings["LoaiSaoLuu"].ToString();
                string v_username = ConfigurationManager.AppSettings["Username"].ToString();
                string v_password = ConfigurationManager.AppSettings["Password"].ToString();
                string v_duongdan_saoluu = ConfigurationManager.AppSettings["DuongDanBackup"].ToString();

                var filename = "";

                string state = "";
                switch (DateTime.Now.DayOfWeek)
                {
                    case DayOfWeek.Friday:
                        state = "full";
                        filename = "backup_full_" + DateTime.Now.ToString("yyyyMMdd");
                        break;
                    default:
                        state = "log";
                        filename = "backup_log_" + DateTime.Now.ToString("yyyyMMdd");
                        break;
                }

                if (!string.IsNullOrEmpty(v_loai_saoluu))
                {
                    state = v_loai_saoluu;
                }

                if (v_loai_csdl.Equals("MSSQL"))
                {
                    repo_SQL.Backup(v_server, v_username, v_password, state, v_ten_csdl, v_duongdan_saoluu, filename); //backup full hoặc log
                }
            }
            catch(Exception ex)
            {
                repo_SQL.WriteLog("Backup NCPT", ex.Message);
            }
        }
    }
}
