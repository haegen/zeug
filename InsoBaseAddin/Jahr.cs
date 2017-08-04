using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InsoBaseAddin
{
    class Jahr
    {
        private int mJahr;
        private decimal mSummeHaben;
        private List<int> mMonate;

        public void SetJahr(int pJahr)
        {
            mJahr = pJahr;
        }

        public int GetJahr()
        {
            return mJahr;
        }

        public void Init(int pMonat)
        {
            mMonate = new List<int>();
            mMonate.Add(pMonat);
        }

        public void SetSummeHaben(decimal pSumHaben)
        {
            mSummeHaben = mSummeHaben + pSumHaben;
        }

        public void AddMonat(int pMonat)
        {
            if (!mMonate.Contains(pMonat))
                mMonate.Add(pMonat);
        }

        public decimal GetUmsatz()
        {
            return mSummeHaben / mMonate.Count;
        }
    }
}
