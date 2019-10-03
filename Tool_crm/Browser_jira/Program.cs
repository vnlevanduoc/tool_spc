using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Browser_jira.Models;
using Newtonsoft.Json;

namespace Browser_jira
{
    class Program
    {
        static void Main(string[] args)
        {
            repo_browser repo = new repo_browser();
            var cultureInfo = CultureInfo.CreateSpecificCulture("vi-VN");
            var date_now = DateTime.Now;
            PhieuYeuCau temp_pyc = new PhieuYeuCau();
            try
            {
                //Get dữ liệu từ file json theo giờ đc chạy
                var src_file_json = @"C:\KHAITHAC_GROUP\Data_CRM.json";
                var jsonData_pyc = File.ReadAllText(src_file_json);
                var list_data_pyc = JsonConvert.DeserializeObject<List<PhieuYeuCau>>(jsonData_pyc).OrderBy(s => DateTime.ParseExact(s.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", cultureInfo)).ToList() ?? new List<PhieuYeuCau>();


                //foreach (var item in list_data_pyc)
                //{
                //    repo.GhiPhieu_ThucHien(item);
                //}

                var item = list_data_pyc.Where(w => w.ThoiGianBatDau == date_now.ToString("dd/MM/yyyy HH:mm") + ":00").ToList();

                if (item.Count() > 0)
                {
                    temp_pyc = item.FirstOrDefault();
                    repo.GhiPhieu_ThucHien(temp_pyc);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                string json_data = JsonConvert.SerializeObject(temp_pyc);
                repo.Send("duoclv.it@evnspc.vn", "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Cần lập phiếu", json_data + "</br>" + ex.Message);
            }
        }
    }
}
