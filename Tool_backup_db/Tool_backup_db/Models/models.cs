namespace Backup_Db_ATD.Models
{
    public class JSON_FILE_BACKUP
    {
        public string TenFile { set; get; }
        public string NgayTao { set; get; }
        public string NgayBatDau { set; get; }
        public string NgayKetThuc { set; get; }
        public string DungLuong { set; get; }
    }

    public class JSON_FILE_DUNGLUONG
    {
        public string HeThong { set; get; }
        public string NgayThucHien { set; get; }
        public string TongDungLuong { set; get; }
        public string DungLuongDuLieu { set; get; }
        public string DungLuongLog { set; get; }
        public string DungLuongKhac { set; get; }
    }
}
