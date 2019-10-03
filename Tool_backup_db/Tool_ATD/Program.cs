using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using Tool_ATD.Models;

namespace Tool_ATD
{
    class Program
    {
        static void Main(string[] args)
        {
            OpenWordUltil openWordUltil = new OpenWordUltil();
            Random rnd = new Random();
            string Source_File_Template = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File\\Template.docx");
            var file_json_path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File/data.json");
            var jsonData = File.ReadAllText(file_json_path);
            var list_data = JsonConvert.DeserializeObject<FILE_JSON>(jsonData) ?? new FILE_JSON();

            file_json_path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File/ngay_nghi_trong_nam.json");
            jsonData = File.ReadAllText(file_json_path);
            var list_ngay_nghi = JsonConvert.DeserializeObject<List<Ngay_Nghi>>(jsonData) ?? new List<Ngay_Nghi>();

            foreach (var item_nhanvien in list_data.NhanVien)
            {
                int i = 0;
                foreach (var item_quy in list_data.Quy)
                {
                    i++;
                    string Soucre_File_Access = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "File\\" + item_nhanvien.NickName + "_ATD_THANG_THU_" + i + ".docx");
                    if (File.Exists(Soucre_File_Access))
                    {
                        File.Delete(Soucre_File_Access);
                    }
                    File.Copy(Source_File_Template, Soucre_File_Access, true);

                    Dictionary<string, string> keyValues = new Dictionary<string, string>();
                    keyValues.Add("V_THANG", item_quy.Thang);
                    keyValues.Add("V_HOTEN", item_nhanvien.HoTen);
                    keyValues.Add("V_CHUYENMON", item_nhanvien.TrinhDo);
                    keyValues.Add("V_CHUCDANHCONGVIEC", item_nhanvien.ChucDanhCongViec);
                    keyValues.Add("V_BACLUONG", item_nhanvien.BacLuong);
                    keyValues.Add("V_NOIDUNGCONGVIEC", item_nhanvien.NoiDungCongViec);
                    keyValues.Add("V_CHUKY", item_nhanvien.HoTen);
                    keyValues.Add("V_TUNGAY", item_quy.NgayBatDau);
                    keyValues.Add("V_DENNGAY", item_quy.NgayKetThuc);

                    DataTable data = new DataTable();
                    data.Columns.Add("STT");
                    data.Columns.Add("NGAYTHUCHIEN");
                    data.Columns.Add("CONGVIECTHUCHIEN");
                    data.Columns.Add("GIOBATDAU");
                    data.Columns.Add("GIOKETTHUC");
                    data.Columns.Add("TONGTHOIGIAN");
                    data.Columns.Add("XACNHANLDP");

                    var startdate = DateTime.ParseExact(item_quy.NgayBatDau, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var enddate = DateTime.ParseExact(item_quy.NgayKetThuc, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var tong_gio_lam_trong_ngay = (Math.Round(int.Parse(item_quy.NgayCong) * float.Parse(item_nhanvien.TyLePhanTram))) * 8 / int.Parse(item_quy.NgayCong);
                    int stt = 0;

                    while (startdate <= enddate)
                    {
                        //loại các ngày nghỉ trong năm
                        var _flag = list_ngay_nghi.Where(s => s._Ngay == startdate);
                        if (_flag.Count() > 0)
                        {
                            startdate = startdate.AddDays(1);
                            continue;
                        }

                        stt++;
                        switch (startdate.DayOfWeek)
                        {
                            case DayOfWeek.Saturday:
                                startdate = startdate.AddDays(2);
                                if (startdate > enddate) //Loại các trường hợp ngày đang chạy vượt quá ngày cuối tháng do rơi vào T7, CN chương trình sẽ công lên
                                {
                                    continue;
                                }
                                break;
                            case DayOfWeek.Sunday:
                                startdate = startdate.AddDays(1);
                                if (startdate > enddate) //Loại các trường hợp ngày đang chạy vượt quá ngày cuối tháng do rơi vào T7, CN chương trình sẽ công lên
                                {
                                    continue;
                                }
                                break;
                            default:
                                break;
                        }

                        var rnd_phut = rnd.Next(0, 59);
                        var ngay_thuc_hien = startdate.ToString("dd/MM/yyyy");
                        var gio_bat_dau = DateTime.ParseExact(ngay_thuc_hien + " 08:" + (rnd_phut < 10 ? "0" + rnd_phut : rnd_phut.ToString()), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                        var gio_nghitrua = DateTime.ParseExact(ngay_thuc_hien + " 12:00", "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                        var gio_bat_dau_buoi_chieu = DateTime.ParseExact(ngay_thuc_hien + " 13:00", "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                        var gio_ket_thuc = gio_bat_dau.AddHours(tong_gio_lam_trong_ngay);
                        if (gio_bat_dau.AddHours(tong_gio_lam_trong_ngay) > gio_nghitrua)
                        {
                            var gio_lam_buoi_sang = gio_nghitrua - gio_bat_dau;

                            var gio_lam_con_lai_buoi_chieu = tong_gio_lam_trong_ngay * 60 - gio_lam_buoi_sang.TotalMinutes;

                            gio_ket_thuc = gio_bat_dau_buoi_chieu.AddMinutes(gio_lam_con_lai_buoi_chieu);
                        }

                        TimeSpan ts = TimeSpan.FromHours(tong_gio_lam_trong_ngay);
                        string txtDate = string.Format("{0}", new DateTime(ts.Ticks).ToString("HH:mm"));

                        data.Rows.Add(new object[]
                        {
                            stt.ToString(),
                            ngay_thuc_hien,
                            item_nhanvien.CongViecThucHien,
                            gio_bat_dau.ToString("HH:mm"),
                            gio_ket_thuc.ToString("HH:mm"),
                            txtDate,
                            ""
                        });
                        startdate = startdate.AddDays(1);
                    }

                    openWordUltil.SetContentControlText(Soucre_File_Access, keyValues);
                    openWordUltil.EditTable(Soucre_File_Access, 0, 2, data);
                }
            }
        }
    }
}
