using System;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace CSharpComLibrary
{
    [Guid("6367D977-2F27-4D09-B25E-9EABF0C534C0"),
InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ComClasEvents
    {
    }
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMyKad
    {
        String[] ICRTextLine { get; set; }
        [DispId(23)]
        String[] GetDetailsFromKAD(String Line);

        [DispId(24)]
        String GetDOB(String Line);

        [DispId(25)]
        String GetGender(String Line);

        [DispId(26)]
        String GetState(String Line);

        [DispId(27)]
        String GetAge(String Line);

    }

    [Guid("C5E38920-97F3-406B-8772-F6A90727A972"),
    ClassInterface(ClassInterfaceType.None),
       ComSourceInterfaces(typeof(ComClasEvents))]
    [ProgId("CSharpComLibrary.MyKad")]
    public class MyKad : IMyKad
    {
        string Dob = ""; string Gender = ""; string State = ""; string Age = "";
        [ComVisible(true)]
        public String[] ICRTextLine { get; set; }
        [ComVisible(true)]
        public String[] GetDetailsFromKAD(String ICRTextLine)
        {
            string[] ICRTextArray = { };
            if (ICRTextLine.ToString().Trim() != "" || ICRTextLine.ToString().Trim() != string.Empty)
            {
                ICRTextArray = ICRTextLine.Split('-');
            }
            return ICRTextArray;
        }

        [ComVisible(true)]
        public String GetDOB(String ICRTextLine)
        {
            string[] list = ICRTextLine.Split('-');
            Dob = list[0];
            string datedob = Dob.Substring(4, 2) + '/' + Dob.Substring(2, 2) + '/' + "19" + Dob.Substring(0, 2);
            return datedob;
        }
        [ComVisible(true)]
        public String GetGender(String ICRTextLine)
        {
            string[] list = ICRTextLine.Split('-');

            Gender = list[2];
            if (Convert.ToInt32(Gender.Trim()) % 2 == 0)
            {
                Gender = "Female";
            }
            else
            {
                Gender = "Male";
            }

            return Gender;
        }
        [ComVisible(true)]
        public String GetState(String ICRTextLine)
        {
            string result="";
            string[] list = ICRTextLine.Split('-');            
            State = list[1];
            #region
            switch (list[1])
            {
                case "01":
                    result ="JOHOR";
                    break;
                case "02":
                    result ="KEDAH";
                        break;
                     case "03":
                    result ="KELANTAN";
                        break;
                     case "04":
                    result ="MELAKA";
                        break;
                     case "05":
                    result ="NEGERI SEMBILAN";
                        break;
                     case "06":
                    result ="PAHANG";
                        break;
                     case "07":
                    result ="PULAU PINANG";
                        break;
                     case "08":                
                    result ="PERAK";
                        break;
                    case "09":
                    result ="PERLIS";
                    break;
                    case "10":
                    result ="SABAH";
                        break;
                     case "11":
                    result ="SARAWAK";
                        break;
                     case "12":
                    result ="SELANGOR";
                        break;
                     case "13":
                    result ="TERENGGANU";
                        break;
                     case "14":
                    result ="WP KUALA LUMPUR";
                        break;
                     case "15":
                    result ="WP LABUAN";
                        break;
                     case "16":                
                    result ="WP PUTRAJAYA";
                        break;
                default:
                    result="Out Of State";
                    break;
            }
            #endregion

            return result;
        }
        [ComVisible(true)]
        public String GetAge(String ICRTextLine)
        {
            string[] list = ICRTextLine.Split('-');
           Dob = list[0];

            string sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;

            DateTime datedob = Convert.ToDateTime(Dob.Substring(2, 2) + '/' + Dob.Substring(4, 2) + '/' + "19" + Dob.Substring(0, 2));
   

            //string x = DateTime.Now.ToString("MM-dd-yyyy");

            //DateTime Now = Convert.ToDateTime(x);
            int Years = new DateTime(DateTime.Now.Subtract(datedob).Ticks).Year - 1;
            //DateTime PastYearDate = datedob.AddYears(Years);
            //int Months = 0;
            //for (int i = 1; i <= 12; i++)
            //{
            //    if (PastYearDate.AddMonths(i) == Now)
            //    {
            //        Months = i;
            //        break;
            //    }
            //    else if (PastYearDate.AddMonths(i) >= Now)
            //    {
            //        Months = i - 1;
            //        break;
            //    }
            //}
            //int Days = Now.Subtract(PastYearDate.AddMonths(Months)).Days;
            //int Hours = Now.Subtract(PastYearDate).Hours;
            //int Minutes = Now.Subtract(PastYearDate).Minutes;
            //int Seconds = Now.Subtract(PastYearDate).Seconds;
            return Years.ToString();

          // return String.Format("Age: {0} Year(s) {1} Month(s) {2} Day(s) {3} Hour(s) {4} Second(s)",
           //  Years, Months, Days, Hours, Seconds);
        }
    }
    //  public class Kadc
    //  {
    //     public static void Main(string[] args)
    //     {
    //         IMyKad kad = new MyKad();
    //        kad.GetDetailsFromKAD("ICRTextLine");

    //        Console.Read();
    //    }

    // }
}


