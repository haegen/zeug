using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsoBaseAddin
{
    class DatumValuePair
    {

        decimal mWert;
        public decimal Wert
        {
            get { return mWert; }
            set { mWert = value; }
        }

        DateTime mDatum;
        public DateTime Datum
        {
            get { return mDatum; }
            set { mDatum = value; }
        }

        int mExcelRow;
        public int ExcelRow
        {
            get { return mExcelRow; }
            set { mExcelRow = value; }
        }


        public DatumValuePair(decimal pWert, DateTime pDatum,int pExcelRow)
        {
            mWert = pWert;
            mDatum = pDatum;
            mExcelRow = pExcelRow;
        }

    }
}
