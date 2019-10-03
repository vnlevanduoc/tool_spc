using System;
using System.Collections.Generic;
namespace Tool_crm.Model
{
    public class FILE_JSON
    {
        public List<NhanVien> NhanVien { set; get; }
        public List<DonVi> DonVi { set; get; }
        public List<IssueType> IssueType { set; get; }
        public List<MoTa_ThacMac> MoTa_ThacMac { set; get; }
        public List<NguoiLienHe> NguoiLienHe { set; get; }
        public List<Project> Project { set; get; }
    }
}
