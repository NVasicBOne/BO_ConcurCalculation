using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CapgeminiConcur
{
    class Parametri
    {
        private static SAPbobsCOM.Company COMPANY = null;
        public static SAPbobsCOM.Company m_GetConnected_COMPANY()
        {
            // DI Company connect:
            if (COMPANY == null)
                COMPANY = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
            if (COMPANY != null && !COMPANY.Connected)
                COMPANY.Connect();
            return COMPANY;
        }
    }
}
