using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    static class StringExtensions
    {
        public static SecureString ToSecureString(this string str)
        {
            SecureString pwd = new SecureString();
            foreach (char c in str)
            {
                pwd.AppendChar(c);
            }

            return pwd;
        }

    }
}
