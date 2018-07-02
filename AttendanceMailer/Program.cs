using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceMailer
{
    static class Program
    {
        static void Main()
        {
            SendMailService s1 = new SendMailService();
            s1.SendmailtoEmployee();
        }
    }
}
