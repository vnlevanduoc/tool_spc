using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using Backup_Db_ATD.Models;
using Newtonsoft.Json;
using WinSCP;

namespace Backup_Db_ATD.Repositorys
{
    public class Repo
    {
        /// <summary>
        /// Lưu lại log backup
        /// </summary>
        /// <param name="path">Đường dẫn lưu file backup file lưu ý ký tự '\' ở cuối phải có (E:\Database\duoclv\)</param>
        /// <param name="_file">Object chứa thông tin file, lưu ý tên file (bảo gồm phần mở rộng), ngày bắt đầu, ngày kết thúc, dung lượng thì thêm vào sau</param>
        public void Ghi_Log_Backup(string path, JSON_FILE_BACKUP _file)
        {
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            var file_extention = _file.TenFile.Split('.').LastOrDefault();

            DirectoryInfo d = new DirectoryInfo(path);
            FileInfo[] Files = d.GetFiles("*." + file_extention);
            
            Console.WriteLine("Thông tin file " + _file.TenFile + ": ");
            var dungluong_file = d.GetFiles(_file.TenFile).FirstOrDefault().Length / (1048576); //đơn vị tính là Mb
            var ngaytao_file = d.GetFiles(_file.TenFile).FirstOrDefault().CreationTime.ToString("dd/MM/yyyy HH:mm");

            var file_json_info = new JSON_FILE_BACKUP()
            {
                TenFile = _file.TenFile,
                NgayTao = ngaytao_file,
                NgayBatDau = _file.NgayBatDau,
                NgayKetThuc = _file.NgayKetThuc,
                DungLuong = dungluong_file.ToString()
            };
            var file_json_path = path + "log_backup.json";

            if (!File.Exists(file_json_path))
            {
                FileStream fs = new FileStream(file_json_path, FileMode.OpenOrCreate);
                StreamWriter str = new StreamWriter(fs);
                str.BaseStream.Seek(0, SeekOrigin.End);
                str.Flush();
                str.Close();
                fs.Close();
            }

            var jsonData = File.ReadAllText(file_json_path); //Đọc lại file
            var list_data = JsonConvert.DeserializeObject<List<JSON_FILE_BACKUP>>(jsonData) ?? new List<JSON_FILE_BACKUP>();
            list_data.Add(file_json_info); //Thêm dữ liệu mới vào
            jsonData = JsonConvert.SerializeObject(list_data.OrderBy(o => DateTime.ParseExact(o.NgayTao, "dd/MM/yyyy HH:mm", cultureInfo)).ToList());
            File.WriteAllText(file_json_path, jsonData);
        }

        public void Upload_FTP(string path_local, string file_extention, string path_server)
        {
            DirectoryInfo d = new DirectoryInfo(path_local);
            FileInfo[] Files = d.GetFiles("*" + file_extention);

            //1. Khởi tạo và đăng nhập
            Console.WriteLine("1. Dang nhap FTP 10.170.3.2.");
            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Ftp,
                HostName = "10.170.3.2",
                UserName = "Khaithac",
                Password = "abcDE@123"
            };
            using (Session session = new Session())
            {

                session.Open(sessionOptions);

                List<FileInfo> list_upload = new List<FileInfo>();
                #region 2. Kiểm tra danh sách ở local và server xem file nào cần được upload
                Console.WriteLine("2. Dang tim file can upload.");
                RemoteDirectoryInfo directory = session.ListDirectory(path_server);
                Files = d.GetFiles("*" + file_extention);
                foreach (FileInfo item_local in Files)
                {
                    var flag = true;
                    foreach (RemoteFileInfo item_server in directory.Files.Where(s => s.FileType != 68))
                    {
                        if (item_local.Name == item_server.Name)
                        {
                            flag = false;
                            break;
                        }
                    }

                    if (flag)
                    {
                        list_upload.Add(item_local);
                    }
                }
                #endregion

                Console.WriteLine("3. Dang xoa file tren server chi giu lai 4 file moi nhat va 1 file chuan bi upload o local");
                int i = 0;
                #region Step 2: Delete old file in server
                directory = session.ListDirectory(path_server); //Lấy lại danh sách file tren server sau khi upload thanh cong
                while (directory.Files.Where(s => s.FileType != 68).Count() > 9)
                {
                    i++;
                    RemoteFileInfo item = directory.Files.Where(s => s.FileType != 68).OrderBy(o => o.LastWriteTime).FirstOrDefault();
                    RemovalOperationResult removeResult = session.RemoveFiles(item.FullName);
                    removeResult.Check();

                    foreach (RemovalEventArgs removeFile in removeResult.Removals)
                    {
                        Console.WriteLine("\t3." + i + ". Xoa file " + removeFile.FileName);
                        WriteLog("Backup_Database", "Step 1: Remove file server: " + removeFile.FileName + " | size: " + item.Length + " bytes");
                    }
                    directory = session.ListDirectory(path_server); //Lấy lại danh sách file trên server
                }
                #endregion

                Console.WriteLine("4. Cac file can upload la: " + (list_upload.Count() > 0 ? list_upload.Select(s => s.Name).Aggregate((a, b) => a + ", " + b) : " "));
                i = 0;
                foreach (var item_upload in list_upload)
                {
                    i++;
                    #region Step 1: Upload file
                    // Upload files
                    TransferOptions transferOptions = new TransferOptions();
                    transferOptions.TransferMode = TransferMode.Binary;

                    TransferOperationResult transferResult;
                    transferResult = session.PutFiles(path_local + item_upload.Name, path_server, false, transferOptions);

                    //3. Đang upload file
                    Console.WriteLine("\t3." + i + ". Đang upload file " + item_upload.Name);
                    transferResult.Check();

                    // Print results
                    foreach (TransferEventArgs transfer in transferResult.Transfers)
                    {
                        Console.WriteLine("\t3." + i + ". Upload file " + transfer.FileName + " thanh cong");
                        WriteLog("Backup_Database", "Step 2: Upload file: " + transfer.FileName + " | size: " + item_upload.Length + " bytes");
                    }
                    #endregion
                }
                Console.WriteLine("6. Da hoan tat.");
            }
        }

        public void MSSQL_Ghi_Log_DungLuong(string hethong, string path_log, string path_database)
        {
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            string v_file_data = ConfigurationManager.AppSettings["MSSQL_File_Database"].ToString();
            string v_file_log = ConfigurationManager.AppSettings["MSSQL_Log_Database"].ToString();

            DriveInfo di = new DriveInfo(path_log.Split(':').FirstOrDefault() + @":\");
            DirectoryInfo d = new DirectoryInfo(path_database);
            var file_db = v_file_data;
            var file_log_db = v_file_log;

            var _size_db = d.GetFiles(file_db).FirstOrDefault().Length / 1048576; //đơn vị tính là Mb
            var _size_log_db = d.GetFiles(file_log_db).FirstOrDefault().Length / 1048576; //đơn vị tính là Mb

            var _size_disk = di.TotalSize / 1048576; //đơn vị tính là Mb
            var _size_disk_free = di.TotalFreeSpace / 1048576; //đơn vị tính là Mb

            //dung lượng khác = _size_disk - (_size_db + _size_log_db + _size_disk_free)

            var file_json_info = new JSON_FILE_DUNGLUONG()
            {
                HeThong = hethong,
                NgayThucHien = DateTime.Now.ToString("dd/MM/yyyy"),
                TongDungLuong = _size_disk.ToString(),
                DungLuongDuLieu = _size_db.ToString(),
                DungLuongLog = _size_log_db.ToString(),
                DungLuongKhac = (_size_disk - (_size_db + _size_log_db + _size_disk_free)).ToString()
            };

            var file_json_path = path_log + "log_size_system.json";

            if (!File.Exists(file_json_path))
            {
                FileStream fs = new FileStream(file_json_path, FileMode.OpenOrCreate);
                StreamWriter str = new StreamWriter(fs);
                str.BaseStream.Seek(0, SeekOrigin.End);
                str.Flush();
                str.Close();
                fs.Close();
            }

            var jsonData = File.ReadAllText(file_json_path); //Đọc lại file
            var list_data = JsonConvert.DeserializeObject<List<JSON_FILE_DUNGLUONG>>(jsonData) ?? new List<JSON_FILE_DUNGLUONG>();
            var _data_temp = list_data.OrderByDescending(o => DateTime.ParseExact(o.NgayThucHien, "dd/MM/yyyy", cultureInfo)).FirstOrDefault(); //Lưu dữ liệu của tháng trước
            list_data.Add(file_json_info); //Thêm dữ liệu mới vào
            jsonData = JsonConvert.SerializeObject(list_data.OrderByDescending(o => DateTime.ParseExact(o.NgayThucHien, "dd/MM/yyyy", cultureInfo)).ToList());
            File.WriteAllText(file_json_path, jsonData);


            #region Ghi file BM.QT.22.21
            string Source_File_Template_BM_QT_22_21 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template/Template_BM_QT_22_21.docx");
            string Soucre_File_Access_BM_QT_22_21 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template/Template_BM_QT_22_21_" + DateTime.Now.ToString("yyyyMMdd") + ".docx");


            if (File.Exists(Soucre_File_Access_BM_QT_22_21))
            {
                File.Delete(Soucre_File_Access_BM_QT_22_21);
            }

            // Create a copy of the template file and open the copy 
            File.Copy(Source_File_Template_BM_QT_22_21, Soucre_File_Access_BM_QT_22_21, true);

            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            keyValues.Add("HeThongDungChung", "Thị trường điện");

            #region Create datatable
            var data = new DataTable();
            data.Columns.Add("TenCSDL");
            data.Columns.Add("ThoiDiemThucHien");
            data.Columns.Add("TongDungLuong");
            data.Columns.Add("DungLuongDuLieu");
            data.Columns.Add("DungLuongLog");
            data.Columns.Add("DungLuongKhac");
            data.Columns.Add("TangTruongDuLieu");
            data.Columns.Add("TangTruongLog");
            data.Columns.Add("TangTruongKhac");
            data.Columns.Add("TangTruongHeThong");
            data.Columns.Add("TangTruongToanHeThong");


            //Dữ liệu tháng trước
            var dulieu_thangtruoc = _data_temp;

            foreach (var item in list_data.OrderByDescending(o => DateTime.ParseExact(o.NgayThucHien, "dd/MM/yyyy", cultureInfo)).ToList())
            {
                int tangtruong_dulieu = dulieu_thangtruoc == null ? 0 : int.Parse(item.DungLuongDuLieu) - int.Parse(dulieu_thangtruoc.DungLuongDuLieu);
                int tangtruong_log = dulieu_thangtruoc == null ? 0 : int.Parse(item.DungLuongLog) - int.Parse(dulieu_thangtruoc.DungLuongLog);
                int tangtruong_khac = dulieu_thangtruoc == null ? 0 : int.Parse(item.DungLuongKhac) - int.Parse(dulieu_thangtruoc.DungLuongKhac);
                int tangtruong_hethong = dulieu_thangtruoc == null ? 0 : tangtruong_dulieu + tangtruong_log + tangtruong_khac;
                data.Rows.Add(new object[] {
                    "NCPT",
                    item.NgayThucHien,
                    item.TongDungLuong,
                    item.DungLuongDuLieu,
                    item.DungLuongLog,
                    item.DungLuongKhac,
                    tangtruong_dulieu,
                    tangtruong_log,
                    tangtruong_khac,
                    tangtruong_hethong,
                    dulieu_thangtruoc == null ? 0 : tangtruong_hethong/(int.Parse(dulieu_thangtruoc.TongDungLuong) + int.Parse(dulieu_thangtruoc.DungLuongLog) + int.Parse(dulieu_thangtruoc.DungLuongKhac))
                });
            }
            #endregion

            OpenWordUltil openWordUltil = new OpenWordUltil();
            openWordUltil.SetContentControlText(Soucre_File_Access_BM_QT_22_21, keyValues);
            openWordUltil.EditTable(Soucre_File_Access_BM_QT_22_21, 0, 3, data);
            #endregion
        }

        public void MSSQL_Ghi_File_BM_QT_22_04(List<JSON_FILE_BACKUP> list_data, string MaSo)
        {
            
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            var thang_thuc_hien = DateTime.ParseExact(list_data.Select(s => s.NgayTao).FirstOrDefault(), "dd/MM/yyyy HH:mm", cultureInfo);

            string Source_File_Template_BM_QT_22_04 = Path.Combine(Environment.CurrentDirectory, "Template/Template_BM_QT_22_04.docx");
            string Soucre_File_Access_BM_QT_22_04 = Path.Combine(Environment.CurrentDirectory, "Template/Template_BM_QT_22_04_Thang" + thang_thuc_hien.ToString("MMyyyy") + ".docx");

            if (File.Exists(Soucre_File_Access_BM_QT_22_04))
            {
                File.Delete(Soucre_File_Access_BM_QT_22_04);
            }

            // Create a copy of the template file and open the copy 
            File.Copy(Source_File_Template_BM_QT_22_04, Soucre_File_Access_BM_QT_22_04, true);

            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            keyValues.Add("MaSo", "QT.22." + MaSo);

            #region Create datatable
            OpenWordUltil openWordUltil = new OpenWordUltil();
            DataTable data = new DataTable();
            data.Columns.Add("NgayGioThucHien");
            data.Columns.Add("HangMuc");
            data.Columns.Add("TinhTrang");
            data.Columns.Add("DienGiai");
            data.Columns.Add("NguoiThucHien");
            data.Columns.Add("NgayKetThuc");

            foreach (var item in list_data.OrderBy(o => DateTime.ParseExact(o.NgayTao, "dd/MM/yyyy HH:mm", cultureInfo)).ToList())
            {
                data.Rows.Add(new object[] { item.NgayKetThuc, "PING", "OK", "", "duoclv", item.NgayKetThuc });
            }
            #endregion

            openWordUltil.SetContentControlText(Soucre_File_Access_BM_QT_22_04, keyValues);
            openWordUltil.EditTable(Soucre_File_Access_BM_QT_22_04, 0, 2, data);
        }

        public void MSSQL_Ghi_File_BM_QT_22_12(List<JSON_FILE_BACKUP> list_data, string url_backup, string ten_du_an)
        {
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            var thang_thuc_hien = DateTime.ParseExact(list_data.Select(s => s.NgayTao).FirstOrDefault(), "dd/MM/yyyy HH:mm", cultureInfo);

            string Source_File_Template_BM_QT_22_12 = Path.Combine(Environment.CurrentDirectory, "Template/Template_BM_QT_22_12.docx");
            string Soucre_File_Access_BM_QT_22_12 = Path.Combine(Environment.CurrentDirectory, "Template/Template_BM_QT_22_12_Thang" + thang_thuc_hien.ToString("MMyyyy") + ".docx");


            if (File.Exists(Soucre_File_Access_BM_QT_22_12))
            {
                File.Delete(Soucre_File_Access_BM_QT_22_12);
            }

            // Create a copy of the template file and open the copy 
            File.Copy(Source_File_Template_BM_QT_22_12, Soucre_File_Access_BM_QT_22_12, true);

            DateTime NgayHienTai = DateTime.Now;
            var NgayDauThang = new DateTime(NgayHienTai.Year, NgayHienTai.Month, 1);
            var NgayCuoiThang = NgayDauThang.AddMonths(1).AddDays(-1);
            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            keyValues.Add("TenDuAn", ten_du_an);
            keyValues.Add("NguoiTao", "duoclv");
            keyValues.Add("TuNgayDenNgay", NgayDauThang.ToString("dd/MM/yyyy") + " - " + NgayCuoiThang.ToString("dd/MM/yyyy"));

            #region Create datatable
            OpenWordUltil openWordUltil = new OpenWordUltil();
            DataTable data = new DataTable();
            data.Columns.Add("ThoiDiemThucHien");
            data.Columns.Add("PhuongThucSaoLuu");
            data.Columns.Add("LoaiSaoLuu");
            data.Columns.Add("GiaiPhapVaCongCuThucHien");
            data.Columns.Add("ViTriSaoLuu");
            data.Columns.Add("TongDungLuong");
            data.Columns.Add("ThoiDiemKetThuc");
            data.Columns.Add("KetQuaThucHien");
            foreach (var item in list_data.OrderBy(o => DateTime.ParseExact(o.NgayTao, "dd/MM/yyyy HH:mm", cultureInfo)).ToList())
            {
                var loai_backup = "Full";
                switch (DateTime.ParseExact(item.NgayTao, "dd/MM/yyyy HH:mm", cultureInfo).DayOfWeek)
                {
                    case DayOfWeek.Friday:
                        loai_backup = "Full";
                        break;
                    default:
                        loai_backup = "Log";
                        break;
                }
                data.Rows.Add(new object[] { item.NgayBatDau, loai_backup, 0, "", url_backup, item.DungLuong, item.NgayKetThuc, "Đạt" });
            }
            #endregion

            openWordUltil.SetContentControlText(Soucre_File_Access_BM_QT_22_12, keyValues);
            openWordUltil.EditTable(Soucre_File_Access_BM_QT_22_12, 0, 2, data);
        }

        public void MSSQL_Ghi_File_BM_QT_22_19(List<JSON_FILE_BACKUP> list_data, string ten_du_an)
        {
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            string Source_File_Template_BM_QT_22_19 = Path.Combine(Environment.CurrentDirectory, "Template/Template_BM_QT_22_19.docx");
            string Soucre_File_Access_BM_QT_22_19 = Path.Combine(Environment.CurrentDirectory, "Template/Template_BM_QT_22_19_" + DateTime.Now.ToString("yyyyMMdd") + ".docx");


            if (File.Exists(Soucre_File_Access_BM_QT_22_19))
            {
                File.Delete(Soucre_File_Access_BM_QT_22_19);
            }

            // Create a copy of the template file and open the copy 
            File.Copy(Source_File_Template_BM_QT_22_19, Soucre_File_Access_BM_QT_22_19, true);
            DateTime NgayHienTai = DateTime.Now;
            var Thang = NgayHienTai.ToString("MM");
            var Nam = NgayHienTai.ToString("yyyy");
            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            keyValues.Add("Thang", Thang);
            keyValues.Add("Nam", Nam);

            #region Create datatable
            OpenWordUltil openWordUltil = new OpenWordUltil();
            DataTable data = new DataTable();
            data.Columns.Add("ThoiDiemKiemTra");
            data.Columns.Add("ToNhom");
            data.Columns.Add("TenCSDL");
            data.Columns.Add("NhanVienVanHanh");
            data.Columns.Add("NgayCuaFileBackup");
            data.Columns.Add("DungLuongCuaFileBackup");
            data.Columns.Add("KetQuaRestore");
            data.Columns.Add("ThucHienQueryDb");
            data.Columns.Add("LyDoLoi");

            data.Rows.Add(new object[] { DateTime.Now.ToString("dd/MM/yyyy HH:mm"), "Khai Thac", ten_du_an, "duoclv", list_data.OrderBy(o => DateTime.ParseExact(o.NgayTao, "dd/MM/yyyy HH:mm", cultureInfo)).ToList().LastOrDefault().NgayTao, list_data.OrderBy(o => DateTime.ParseExact(o.NgayTao, "dd/MM/yyyy HH:mm", cultureInfo)).ToList().LastOrDefault().DungLuong + " Mb", "Đạt", "Đạt", "" });
            #endregion

            openWordUltil.SetContentControlText(Soucre_File_Access_BM_QT_22_19, keyValues);
            openWordUltil.EditTable(Soucre_File_Access_BM_QT_22_19, 0, 2, data);
        }

        public void WriteLog(string app, string Message)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(@"C:\SPCIT_Log\Log_Backup_NCPT.txt", true);
                sw.WriteLine("[" + app + "][" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "] " + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
    }
}
