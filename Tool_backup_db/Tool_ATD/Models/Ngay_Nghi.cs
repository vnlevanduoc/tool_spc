using System;
using System.Globalization;

namespace Tool_ATD.Models
{
    public class Ngay_Nghi
    {
        public string Ngay;
        public DateTime _Ngay
        {
            get
            {
                return DateTime.ParseExact(Ngay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
        }
    }
}
