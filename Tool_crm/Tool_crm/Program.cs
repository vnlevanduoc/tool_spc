using Microsoft.Win32.TaskScheduler;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tool_crm.Model;

namespace Tool_crm
{
    class Program
    {
        static void Main(string[] args)
        {
            //Kiểm tra xem ngày này có phải ngày đầu tháng ko
            //Nếu phải thì thực hiện random file JIRA
            DateTime date = DateTime.Now;
            #region Khai báo biến global và tạo file
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            var file_json_data = @"C:\KHAITHAC_GROUP\Data_CRM.json";
            if (!File.Exists(file_json_data))
            {
                FileStream fs = new FileStream(file_json_data, FileMode.OpenOrCreate);
                StreamWriter str = new StreamWriter(fs);
                str.BaseStream.Seek(0, SeekOrigin.End);
                str.Flush();
                str.Close();
                fs.Close();
            }
            #endregion

            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            bool chay_lai = bool.Parse(ConfigurationManager.AppSettings["ChayLai"]);

            if (!chay_lai)
            {
                //Kiểm tra xem tháng hiện tại có dữ liệu chưa, nếu chưa thì chay_lai = true
                var thang_hientai = DateTime.Now.Month;
                try
                {
                    var src_file_json = @"C:\KHAITHAC_GROUP\Data_CRM.json";
                    var data_pyc = JsonConvert.DeserializeObject<List<PhieuYeuCau>>(File.ReadAllText(src_file_json)).Where(s => DateTime.ParseExact(s.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", cultureInfo).Month == thang_hientai).ToList();

                    if (data_pyc.Count == 0)
                    {
                        chay_lai = true;
                    }
                }
                catch
                {
                    chay_lai = true;
                }
            }

            if (date.Day == 1 || chay_lai)
            {
                Random rnd = new Random();
                //Tạo file random
                double SoGio = double.Parse(ConfigurationManager.AppSettings["Gio_CRM"]);
                var SoPhieu = Math.Round(SoGio * 60 / 20);
                double SoPhieu_TrongNgay = 0;
                double TongSoGio_ThucHien = 0;
                var file_json_path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File/data.json");
                var jsonData_file = File.ReadAllText(file_json_path);
                var list_data_file = JsonConvert.DeserializeObject<FILE_JSON>(jsonData_file) ?? new FILE_JSON();

                //Thời gian bắt đầu và kết thúc của 1 tháng
                var start_date = DateTime.ParseExact("01/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("yyyy"), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var end_date = DateTime.ParseExact("01/" + DateTime.Now.AddMonths(1).ToString("MM") + "/" + DateTime.Now.AddMonths(1).ToString("yyyy"), "dd/MM/yyyy", CultureInfo.InvariantCulture);

                #region Tạo thông tin phiếu
                SoPhieu_TrongNgay = Math.Ceiling(SoPhieu / GetBusinessDays(start_date, end_date.AddDays(-1)));
                string _nhanvien = ConfigurationManager.AppSettings["NhanVien"];
                int nv = 0; //id =1 là duoclv, 2 là hoapt
                switch (_nhanvien)
                {
                    case "duoclv":
                        nv = 1;
                        break;
                    case "hoapt":
                        nv = 2;
                        break;
                    default:
                        nv = 1;
                        break;

                }
                var nhanvien = list_data_file.NhanVien.Where(s => s.Id == nv).FirstOrDefault();
                var data_pyc = new List<PhieuYeuCau>();
                while (start_date < end_date)
                {
                    switch (start_date.DayOfWeek)
                    {
                        case DayOfWeek.Saturday:
                            start_date = start_date.AddDays(2);
                            break;
                        case DayOfWeek.Sunday:
                            start_date = start_date.AddDays(1);
                            break;
                        default:
                            break;
                    }

                    int solan_loop = int.Parse(SoPhieu_TrongNgay.ToString());
                    if (TongSoGio_ThucHien >= SoGio * 0.87)
                    {
                        solan_loop = rnd.Next(1, int.Parse(SoPhieu_TrongNgay.ToString()));
                    }

                    //Từ số phiếu quy ra đc 1 ngày làm bao nhiêu phiếu
                    for (int i = 0; i < solan_loop; i++)
                    {
                        //Thời gian bắt
                        TimeSpan newSpan = new TimeSpan(rnd.Next(8, 16), rnd.Next(0, 59), 0);
                        DateTime newDate = start_date + newSpan;
                        var gio_bat_dau_nghi_trua = DateTime.ParseExact(newDate.ToString("dd/MM/yyyy") + " 12:00:00", "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                        var gio_ket_thuc_nghi_trua = DateTime.ParseExact(newDate.ToString("dd/MM/yyyy") + " 13:00:00", "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                        if (newDate > gio_bat_dau_nghi_trua && newDate < gio_ket_thuc_nghi_trua)
                        {
                            newSpan = new TimeSpan(rnd.Next(13, 16), rnd.Next(0, 59), 0);
                            newDate = start_date + newSpan;
                        }

                        var ngay_tiep_nhan = newDate;
                        var ngay_thuc_hien = newDate;
                        var ngay_ket_thuc = newDate.AddMinutes(20);
                        if (ngay_ket_thuc >= gio_bat_dau_nghi_trua && ngay_ket_thuc <= gio_ket_thuc_nghi_trua)
                        {
                            break;
                        }
                        TongSoGio_ThucHien += (ngay_ket_thuc - ngay_thuc_hien).TotalHours;

                        //Project
                        //var temp_project = list_data.Project[rnd.Next(list_data.Project.Count)];
                        //var project = temp_project.Name;
                        var issuetype = list_data_file.IssueType.Where(w => w.Id == 4).FirstOrDefault().Name;
                        //var temp_summary_project = list_data.MoTa_ThacMac.Where(w => w.HeThong == temp_project.Id).ToList();

                        var temp_summary = list_data_file.MoTa_ThacMac[rnd.Next(list_data_file.MoTa_ThacMac.Count)];
                        var summary = temp_summary.MoTa;

                        //Từ thắc mắc suy ngược ra project
                        var project = list_data_file.Project.Where(w => w.Id == temp_summary.HeThong).FirstOrDefault().Name;

                        var temp_nguoiyeucau_donvi = list_data_file.NguoiLienHe[rnd.Next(list_data_file.NguoiLienHe.Count)]; //Người yêu cầu có đơn vị rồi nên ko cần random đơn vị
                        var donvi = temp_nguoiyeucau_donvi.DonVi;
                        var nguoiyeucau = temp_nguoiyeucau_donvi.HoTen;
                        var motayeucau = summary;
                        var nguoithuchien = nhanvien.Username;
                        var cachxuly = temp_summary.CachGiaiQuyet;

                        PhieuYeuCau pyc = new PhieuYeuCau()
                        {
                            ThoiGianBatDau = ngay_thuc_hien.ToString("dd/MM/yyyy HH:mm:ss"),
                            ThoiGianKetThuc = ngay_ket_thuc.ToString("dd/MM/yyyy HH:mm:ss"),
                            Project = project,
                            IssueType = issuetype,
                            Summary = summary,
                            DonVi = donvi.ToString(),
                            NguoiYeuCau = nguoiyeucau,
                            MoTaYeuCau = motayeucau,
                            NguoiThucHien = nguoithuchien,
                            CachGiaiQuyet = cachxuly
                        };
                        data_pyc.Add(pyc);
                    }

                    start_date = start_date.AddDays(1);
                }

                //Ghi file
                jsonData_file = JsonConvert.SerializeObject(data_pyc.OrderBy(o => DateTime.ParseExact(o.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", cultureInfo)).ToList());
                File.WriteAllText(file_json_data, jsonData_file);
                #endregion
            }

            //Tiếp theo là tự động đặt tạo các job hằng ngày theo file đã random ở C:\KHAITHAC_GROUP\Duoclv\xxx.json
            //Lấy danh sách phiếu trong ngày hiện tại của file random
            string src_app = ConfigurationManager.AppSettings["SOURCE_APP"];
            #region Tạo file JSON thông tin phiếu của tháng
            var jsonData_pyc = File.ReadAllText(file_json_data);
            var list_data_pyc = JsonConvert.DeserializeObject<List<PhieuYeuCau>>(jsonData_pyc) ?? new List<PhieuYeuCau>();
            //Xóa hết tất cả Task CRM duoclv đang có
            foreach (var item in list_data_pyc)
            {
                var temp_date = DateTime.ParseExact(item.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", cultureInfo).ToString("yyyyMMddHHmmss");
                var taskname = "CRM_" + temp_date;
                using (TaskService ts = new TaskService())
                {
                    try
                    {
                        ts.RootFolder.DeleteTask(taskname);
                        Console.WriteLine("Da xoa " + taskname);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Khong tim thay " + taskname);
                    }
                }
            }

            //Từ danh sách random lấy ra phiếu trong ngày hiện tại và tạo ra các Task tương ứng
            list_data_pyc = list_data_pyc.Where(w => DateTime.ParseExact(w.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", cultureInfo).ToString("dd/MM/yyyy") == DateTime.Now.ToString("dd/MM/yyyy")).ToList();
            foreach (var item in list_data_pyc)
            {
                var temp_date = DateTime.ParseExact(item.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", cultureInfo);

                using (TaskService ts = new TaskService())
                {
                    // Create a new task definition and assign properties
                    TaskDefinition td = ts.NewTask();
                    td.RegistrationInfo.Description = "Chạy CRM";

                    td.Principal.LogonType = TaskLogonType.S4U;
                    td.Settings.Hidden = true;
                    td.Principal.RunLevel = TaskRunLevel.Highest;

                    // Create a trigger that will fire the task at this time every other day
                    TimeTrigger tTrigger = new TimeTrigger();
                    tTrigger.StartBoundary = temp_date; //yyyy,mm,dd,hh,mm,ss
                    td.Triggers.Add(tTrigger);

                    // Create an action that will launch Notepad whenever the trigger fires
                    td.Actions.Add(new ExecAction(src_app, null, null));

                    // Register the task in the root folder
                    ts.RootFolder.RegisterTaskDefinition(@"CRM_" + temp_date.ToString("yyyyMMddHHmmss"), td);
                }
            }
            #endregion
        }

        public static double GetBusinessDays(DateTime startD, DateTime endD)
        {
            double calcBusinessDays = 1 + ((endD - startD).TotalDays * 5 - (startD.DayOfWeek - endD.DayOfWeek) * 2) / 7;
            if (endD.DayOfWeek == DayOfWeek.Saturday) calcBusinessDays--;
            if (startD.DayOfWeek == DayOfWeek.Sunday) calcBusinessDays--;

            return calcBusinessDays;
        }
    }
}
