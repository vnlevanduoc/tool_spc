using Browser_jira.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;

namespace Browser_jira
{
    class repo_browser
    {
        string link_issue = "";
        string error = "";

        public void GhiPhieu_ThucHien(PhieuYeuCau pyc)
        {
            ChromeDriver chromeDriver = new ChromeDriver();
            Console.WriteLine("DANG NHAP");
            #region Đăng nhập
            chromeDriver.Url = "http://10.170.2.70:8081/login.jsp";
            chromeDriver.Navigate();
            DangNhap(chromeDriver, pyc.NguoiThucHien, pyc.NguoiThucHien);
            #endregion

            Console.WriteLine("TAO PHIEU");
            TaoPhieu(chromeDriver, pyc);

            Console.WriteLine("BAT DAU XU LY");
            BatDauXuLy(chromeDriver);

            #region Chuyển đến yêu cần cần hoàn tất
            chromeDriver.Url = link_issue;
            chromeDriver.Navigate();
            #endregion

            Console.WriteLine("DANG GIAI QUYET");
            GiaiQuyetPhieu(chromeDriver, pyc);

            Console.WriteLine("DA DONG");
        }

        public void DangNhap(ChromeDriver chromeDriver, string p_username, string p_password)
        {
            try
            {
                var username = chromeDriver.FindElementById("login-form-username");
                var password = chromeDriver.FindElementById("login-form-password");
                var submit_login = chromeDriver.FindElementById("login-form-submit");
                username.Clear();
                password.Clear();
                username.SendKeys(p_username);
                password.SendKeys(p_password);
                submit_login.Click();
            }
            catch (Exception ex)
            {
                error += "- Lỗi đăng nhập [" + ex.Message + "] - ";
                throw ex;
            }
        }

        public void TaoPhieu(ChromeDriver chromeDriver, PhieuYeuCau pyc)
        {
            try
            {
                //Nhấn nút tạo phiếu
                var btn_create_crm = chromeDriver.FindElementById("create_link");
                btn_create_crm.Click();
                Thread.Sleep(5000);

                //Nhập project
                EditElement(chromeDriver, "project-field", pyc.Project);

                //Nhập loại yêu cầu
                EditElement(chromeDriver, "issuetype-field", pyc.IssueType);

                //Nhập tiêu đề
                EditElement(chromeDriver, "summary", pyc.Summary);

                //Chọn đơn vị
                SelectElement customfield_10007 = new SelectElement(chromeDriver.FindElementById("customfield_10007"));
                customfield_10007.SelectByValue(pyc.DonVi);

                //Nhập thông tin liên hệ
                EditElement(chromeDriver, "customfield_10008", pyc.NguoiYeuCau);

                //Nhập mô tả chi tiết
                EditElement(chromeDriver, "description", pyc.MoTaYeuCau);

                //Chọn mức độ khẩn của yêu cầu
                //EditElement(chromeDriver, "priority-field", "Trung bình");

                //Chọn người thực hiện assign-to-me-trigger
                ClickElement(chromeDriver, "id", "assign-to-me-trigger");

                //Nhấn gửi yêu cầu
                ClickElement(chromeDriver, "id", "create-issue-submit");
                Thread.Sleep(3000); //Chờ 1,5s để tiếp tục
            }
            catch (Exception ex)
            {
                error += "Lỗi tạo phiếu [" + ex.Message + "] - ";
                throw ex;
            }
        }

        public void BatDauXuLy(ChromeDriver chromeDriver)
        {
            try
            {
                //Lưu lại đường dẫn của yêu cầu để tiến hành xử lý
                var html_key_issue = chromeDriver.FindElementByClassName("issue-link");
                var key = html_key_issue.GetAttribute("data-issue-key");
                link_issue = "http://10.170.2.70:8081/browse/" + key;
                chromeDriver.Url = link_issue;
                chromeDriver.Navigate();

                //Kiểm tra xem các nút hiện của của thanh xử lý
                //action_id_301: Đã bắt đầu xử lý
                //action_id_4: Chưa xử lý
                //action_id_5: Nhập nội dung xử lý
                //action_id_2: Đóng yêu cầu
                //Nếu hiện tại có nút action_id_301 thì ko thực hiện bước này mà chờ sang bước xử lý
                var list_element = chromeDriver.FindElementById("opsbar-opsbar-transitions").FindElements(By.ClassName("toolbar-item")).Select(s => s.FindElement(By.TagName("a"))).Select(s => s.GetAttribute("id"));

                if (!list_element.FirstOrDefault().Equals("action_id_301"))
                {
                    //click nút bắt đầu xử lý
                    ClickElement(chromeDriver, "id", "action_id_5");
                    Thread.Sleep(3000); //Chờ 1,5s để tiếp tục
                }
            }
            catch (Exception ex)
            {
                error += "Lỗi bắt đầu xử lý [" + ex.Message + "] - ";
                throw ex;
            }
        }

        public void GiaiQuyetPhieu(ChromeDriver chromeDriver, PhieuYeuCau pyc)
        {
            try
            {
                //click nút bắt đầu giải quyết vấn đề
                ClickElement(chromeDriver, "id", "action_id_5");
                Thread.Sleep(3000);

                //NHẬP CÁC NỘI DUNG GIẢI QUYẾT VẤN ĐỀ
                SelectElement selector_resolution = new SelectElement(chromeDriver.FindElement(By.Id("resolution")));
                selector_resolution.SelectByValue("10400");

                //Nội dung xử lý
                EditElement(chromeDriver, "customfield_10302", pyc.CachGiaiQuyet);

                //Ngày xử lý
                //EditElement(chromeDriver, "customfield_10700", DateTime.ParseExact(pyc.ThoiGianBatDau, "dd/MM/yyyy HH:mm:ss", CultureInfo.CreateSpecificCulture("vi-VN")).ToString("dd/MM/yyyy HH:mm"));

                //Click nút đã giải quyết
                Thread.Sleep(2000);
                ClickElement(chromeDriver, "id", "issue-workflow-transition-submit");

                //Đăng nhập lại bằng user hoapt để đóng phiếu
                chromeDriver.Quit();
            }
            catch (Exception ex)
            {
                error += "Lỗi giải quyết phiếu [" + ex.Message + "] - ";
                throw ex;
            }
        }

        public void DongPhieu(ChromeDriver chromeDriver) //Dùng user hoapt để đóng
        {
            try
            {
                //Kiểm tra xem có nút đóng phiếu không nếu không thì thoát luôn nếu có thì tiếp tục
                var check_button = chromeDriver.FindElements(By.Id("action_id_701")).Count();
                if (check_button > 0)
                {
                    //Nhấn nút bắt đầu đóng issue
                    ClickElement(chromeDriver, "id", "action_id_701");
                    Thread.Sleep(3000);

                    //Nhập người đang đóng phiếu
                    EditElement(chromeDriver, "customfield_10017", "hoapt.it");
                    Thread.Sleep(3000);

                    //Nhấn nút đóng phiếu
                    ClickElement(chromeDriver, "id", "issue-workflow-transition-submit");
                    Thread.Sleep(3000);

                    //Thoát ứng dụng hiện tại để đăng nhập lại
                    chromeDriver.Quit();
                }
            }
            catch (Exception ex)
            {
                chromeDriver.Quit();
                error += "Lỗi đóng phiếu [" + ex.Message + "] - ";
                throw ex;
            }
        }

        public void CheckInputIsDisable(ChromeDriver chromeDriver, string element, string attr)
        {
            var check_is_disable = chromeDriver.FindElementById(element).GetAttribute(attr);
            while (!string.IsNullOrEmpty(check_is_disable))
            {
                Thread.Sleep(8888); //Mỗi lần while thì chờ 1s
                try
                {
                    check_is_disable = chromeDriver.FindElementById(element).GetAttribute(attr);
                }
                catch (StaleElementReferenceException e)
                {
                    check_is_disable = null;
                }
            }
        }

        public void CheckDialogExist(ChromeDriver chromeDriver, string element)
        {
            var staleElement = true;
            while (staleElement)
            {
                Thread.Sleep(8888);
                try
                {
                    var check_dialog = chromeDriver.FindElements(By.Id(element)).Count();
                    if (check_dialog > 0)
                    {
                        staleElement = false;
                    }
                }
                catch
                {
                    staleElement = true;
                }
            }
        }

        public void ClickElement(ChromeDriver chromeDriver, string id_class, string element)
        {
            var flag = false;
            var flag_count = 0;
            while (!flag)
            {
                flag_count++;
                try
                {
                    switch (id_class)
                    {
                        case "id":
                            var id_element = chromeDriver.FindElementById(element);
                            id_element.Click();
                            break;
                        case "class":
                            var class_element = chromeDriver.FindElementByClassName(element);
                            class_element.Click();
                            break;
                        default:
                            break;
                    }
                    flag = true;
                    Thread.Sleep(3000);
                }
                catch (Exception ex)
                {
                    if (flag_count > 5)
                    {
                        throw ex;
                    }
                    Thread.Sleep(5000); //Nếu lỗi thì dừng lại 1s để thử lại
                }
            }
        }

        public void EditElement(ChromeDriver chromeDriver, string element, string connent_txt)
        {
            var flag = false;
            while (!flag)
            {
                try
                {
                    var text_field = chromeDriver.FindElementById(element);
                    text_field.Clear();
                    text_field.SendKeys(connent_txt);
                    text_field.SendKeys(Keys.Enter);
                    flag = true;
                    Thread.Sleep(3000);
                }
                catch
                {
                    Thread.Sleep(3000); //Nếu lỗi thì dừng lại 1s để thử lại
                }
            }
        }

        #region Send Email Code Function
        public void Send(string ToEmail, string Subj, string Message)
        {
            #region cấu hình gửi mail
            string Server = "mail.evnspc.net";
            int Port = 25;
            string Username = "duoclv.it@evnspc.net";
            string Password = "123@abc";
            string Subject = "[" + DateTime.Now.ToString("dd/MM/yyyy HH:mm") + "] Cần lập phiếu";
            #endregion
            try
            {
                MailMessage msg = new MailMessage();
                msg.From = new MailAddress(Username);
                msg.Subject = Subj;
                msg.Body = Message;
                msg.IsBodyHtml = true;

                //Gửi cùng lúc nhiều mail phân biệt bằng dấu ","
                string[] ToMuliId = ToEmail.Split(',');
                foreach (string ToEMailId in ToMuliId)
                {
                    msg.To.Add(new MailAddress(ToEMailId));
                }

                SmtpClient client = new SmtpClient(Server);
                client.Port = Port;
                client.EnableSsl = false;
                client.Timeout = 100000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(Username, Password);
                client.Send(msg);
            }
            catch (Exception ex)
            {
                WriteLog("Lỗi: " + ex.Message);
            }
        }
        #endregion

        public void WriteLog(string Message)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(@"C:\KHAITHAC_GROUP\Duoclv\Log_CRM.txt", true);
                sw.WriteLine("[Success][" + DateTime.Now.ToString() + "] " + Message);
                sw.Flush();
                sw.Close();
            }
            catch
            {
            }
        }
    }
}
