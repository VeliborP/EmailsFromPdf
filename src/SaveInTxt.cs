using System;
using System.IO;

namespace EmailFromPDF
{
    public class SaveInTxt
    {
        public SaveInTxt(string[] tekst, string putanja)
        {
            var date = DateTime.Now.ToString("yy-MM-dd HH_mm_ss");
            File.WriteAllLines(putanja + "\\EmailsTxt" + date + ".txt", tekst);
        }
    }
}
