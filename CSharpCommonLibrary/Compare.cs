//using iTextSharp.text.pdf;
//using iTextSharp.text.pdf.parser;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace CSharpComLibrary
{
    [Guid("7BD20046-DF8C-44A6-8F6B-687FAA26FA71"),
        InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ComClass1Events
    {
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface INeilTest
    {
        String[] branches { get; set; }
        String[] words { get; set; }
        String branchesText { get; set; }

        System.Collections.Generic.List<String> branchesArray { get; set; }
        CustomWordType[] customWords { get; set; }

        [DispId(1)]
        Int32 CheckSubString(String locatorAlternativeValue, String ocrTextLine, Int32 locatorAlternativeIndex, Int32 ocrLine);

        [DispId(2)]
        String ValidatePO(String POText);

        [DispId(2)]
        String GetSubstring(String source, String subString);

        [DispId(3)]
        Boolean HasSubstring(String source, String subString);

        [DispId(4)]
        String GetSubstring(String source, String subString, Boolean isFirst);

        [DispId(5)]
        String OCRLineToTableCell(String OCRTextLine, int getWordAtIndex);

        [DispId(6)]
        String GetWord(int Index);

        //[DispId(7)]
        //String GetTextFromPDF(String path);

       [DispId(7)]
        Int32 GetWordCount(String phrase);

        [DispId(8)]
        Boolean IsCountryExist(string country);

        [DispId(9)]
        String GetBranch(String keyword);

        [DispId(10)]
        CustomWordType[] SortWordByLeftPointer();

        [DispId(11)]
        void Init(long count);

        [DispId(12)]
        void AddCustomWord(CustomWordType addword);

        [DispId(13)]
        CustomWordType GetCustomWordByIndex(long Index);

        [DispId(14)]
        Boolean IsAlpha(String input);

        [DispId(15)]
        Boolean IsAlphaNumeric(String input);

        [DispId(16)]
        Boolean IsNumeric(String input);

        [DispId(17)]
        Boolean HasSpecialChar(String input);

        [DispId(18)]
        String TakeAlpha(String input);

        [DispId(19)]
        String TakeNumeric(String input);

        [DispId(20)]
        String  RemoveSpecialChar(String input);

        [DispId(21)]
        String TakePhraseWithoutSpecialChar(String input);

        [DispId(22)]
       Boolean TestRegex(String input);

      
     
    }

    [Guid("0D53A3E8-E51A-49C7-944E-E72A2064F938"),
        ClassInterface(ClassInterfaceType.None),
        ComSourceInterfaces(typeof(ComClass1Events))]
    [ProgId("CSharpComLibrary.Compare")]
    public class Compare : INeilTest
    {
        [ComVisible(true)]
        public String[] words { get; set; }

        [ComVisible(true)]
        public String[] branches { get; set; }

        [ComVisible(true)]
        public String branchesText { get; set; }

        [ComVisible(true)]
        public System.Collections.Generic.List<String> branchesArray { get; set; }

        public CustomWordType[] customWords { get; set; }

        //CustomWordType[] INeilTest.customWords { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        [ComVisible(true)]
        public Int32 CheckSubString(String locatorAlternativeValue, String ocrTextLine, Int32 locatorAlternativeIndex, Int32 ocrLine)
        {
            string[] ocrTextWords = ocrTextLine.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] locatorAlternativeWords = locatorAlternativeValue.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            Int32 currentocatorIndex = locatorAlternativeIndex;

            Boolean isFound = false;
            if (locatorAlternativeIndex >= locatorAlternativeWords.Length)
                return 9999;

            for (int j = 0; j < ocrTextWords.Length; j++)
            {
                if (locatorAlternativeIndex >= locatorAlternativeWords.Length)
                    return 9999;
                if (ocrTextWords[j] == locatorAlternativeWords[locatorAlternativeIndex])
                {
                    isFound = true;
                    locatorAlternativeIndex++;
                }
                else if (isFound)
                {
                    if (ocrLine > 1)
                        locatorAlternativeIndex = -1;
                    isFound = false;
                    break;
                }
            }
            if (locatorAlternativeIndex == currentocatorIndex)
                return -1;
            return locatorAlternativeIndex;
        }

        [ComVisible(true)]
        public String ValidatePO(String POText)
        {
            string[] POTextArray = POText.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
            if (POTextArray.Length > 2)
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(POTextArray[2], @"^[a-zA-Z]+$"))
                    return POTextArray[0] + "-" + POTextArray[1] + "-" + POTextArray[2].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                else
                    return POTextArray[0] + "-" + POTextArray[1];
            }
            return POText;
        }

        [ComVisible(true)]
        public String GetSubstring(String source, String subString)
        {
            if (source.Contains(subString))
            {
                var sourceArray = source.Split(new String[] { subString }, StringSplitOptions.RemoveEmptyEntries);
                if (sourceArray.Length > 0)
                    return sourceArray[sourceArray.Length - 1];
                else
                    return "";
            }
            return "";
        }

        [ComVisible(true)]
        public String GetSubstring(String source, String subString, Boolean isFirst)
        {
            if (source.Contains(subString))
            {
                var sourceArray = source.Split(new String[] { subString }, StringSplitOptions.RemoveEmptyEntries);
                if (sourceArray.Length > 0)
                {
                    if (!isFirst)
                        return sourceArray[sourceArray.Length - 1];
                    return sourceArray[0];

                }
                else
                    return "";
            }
            return "";
        }

        [ComVisible(true)]
        public Boolean HasSubstring(String source, String subString)
        {
            return source.Contains(subString);
        }

        [ComVisible(true)]
        public String OCRLineToTableCell(String OCRTextLine, int getWordAtIndex)
        {
            string[] textArray = OCRTextLine.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] newTextArray = new string[50];

            int j = 0;
            for (int i = 0; i < textArray.Length - 1; i++)
            {
                string addWord = textArray[i];
                if (textArray[i].Length == 1)
                {
                    if (IsAlphaNumeric(textArray[i]))
                        goto AddWord;
                    if (i == 0)
                        goto AddWord;

                    addWord = textArray[i - 1] + addWord;

                    if (textArray.Length > (i + 1) && textArray[i].Length == 1)
                    {
                        addWord = addWord + textArray[i + 1];
                        i++;
                    }
                    j = j > 0 ? j - 1 : j;
                    //i++;

                }
                else if (!IsAlphaNumeric(textArray[i][textArray[i].Length - 1].ToString()))
                {
                    if (textArray[i][textArray[i].Length - 1] == '$')
                        goto AddWord;

                    addWord = addWord + textArray[i + 1];
                    j = j > 0 ? j - 1 : j;
                    i++;
                }

            AddWord:
                newTextArray[j] = addWord;
                j++;
            }
            words = newTextArray;
            return newTextArray[getWordAtIndex];
        }
        


        [ComVisible(true)]
        public String GetWord(int Index)
        {
            if (words == null)
                return "";
            return words[Index];
        }

        [ComVisible(true)]
        public Int32 GetWordCount(String phrase)
        {
            if (String.IsNullOrEmpty(phrase))
                return 0;
            return phrase.Split(new String[] { " " }, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        //[ComVisible(true)]   
        //public String GetTextFromPDF(String path)
        //{
        //    var text = new System.Text.StringBuilder();
        //    using (PdfReader reader = new PdfReader(@"" + path))
        //    {
        //        for (int i = 1; i <= reader.NumberOfPages; i++)
        //        {
        //            text.Append(PdfTextExtractor.GetTextFromPage(reader, i));

        //        }
        //    }

        //    return text.ToString();
        //}

        [ComVisible(true)]
        public Boolean IsCountryExist(string country)
        {
            string countries = "Afghanistan, Albania, Algeria, Andorra, Angola, Antigua and Barbuda, Argentina, Armenia, Australia, Austria, Austrian Empire, Azerbaijan, Baden*, Bahrain, Bangladesh, Barbados, Bavaria*, Belarus, Belgium, Belize, Benin (Dahomey), Bolivia, Bosnia and Herzegovina, Botswana, Brazil, Brunei, Brunswick and Lüneburg, Bulgaria, Burkina Faso (Upper Volta), Burma, Burundi, Cabo Verde, Cambodia, Cameroon, Canada, Central African Republic, Central American Federation*, Chad, Chile, China, Colombia, Comoros, Costa Rica, Cote d’Ivoire (Ivory Coast), Croatia, Cuba, Cyprus, Czechia, Czechoslovakia, Democratic Republic of the Congo, Denmark, Djibouti, Dominica, Dominican Republic, East Germany (German Democratic Republic), Ecuador, Egypt, El Salvador, Equatorial Guinea, Eritrea, Estonia, Eswatini, Ethiopia, Federal Government of Germany (1848-49)*, Fiji, Finland, France, Gabon, Georgia, Germany, Ghana, Greece, Grenada, Guatemala, Guinea, Guinea-Bissau, Guyana, Haiti , Hanover*, Hanseatic Republics*, Hawaii*, Hesse*, Holy See, Honduras, Hungary, Iceland, India, Indonesia, Iran, Iraq, Ireland, Israel, Italy, Jamaica, Japan, Jordan, Kazakhstan, Kenya, Kingdom of Serbia/Yugoslavia*, Kiribati, Korea, Kosovo, Kuwait, Kyrgyzstan, Laos, Latvia, Lebanon, Lesotho, Lew Chew (Loochoo)*, Liberia, Libya, Liechtenstein, Lithuania, Luxembourg, Macedonia, Madagascar, Malawi, Malaysia, Maldives, Mali, Malta, Marshall Islands, Mauritania, Mauritius, Mecklenburg-Schwerin*, Mecklenburg-Strelitz*, Mexico, Micronesia, Moldova, Monaco, Mongolia, Montenegro, Morocco, Mozambique, Namibia, Nassau*, Nauru, Nepal, New Zealand, Nicaragua, Niger, Nigeria, North German Confederation*, North German Union*, Norway, Oldenburg*, Oman, Orange Free State*, Pakistan, Palau, Panama, Papal States*, Papua New Guinea, Paraguay, Peru, Philippines, Piedmont-Sardinia*, Poland, Portugal, Qatar, Republic of Genoa*, Republic of Korea (South Korea), Republic of the Congo, Romania, Russia, Rwanda, Saint Kitts and Nevis, Saint Lucia, Saint Vincent and the Grenadines, Samoa, San Marino, Sao Tome and Principe, Saudi Arabia, Schaumburg-Lippe*, Senegal, Serbia, Seychelles, Sierra Leone, Singapore, Slovakia, Slovenia, Somalia, South Africa, South Sudan, Spain, Sri Lanka, Sudan, Suriname, Sweden, Switzerland, Syria, Tajikistan, Tanzania, Texas*, Thailand, The Bahamas, The Cayman Islands, The Congo Free State, The Duchy of Parma*, The Gambia, The Grand Duchy of Tuscany*, The Netherlands, The Solomon Islands , The United Arab Emirates, The United Kingdom, Timor-Leste, Togo, Tonga, Trinidad and Tobago, Tunisia, Turkey, Turkmenistan, Tuvalu, Two Sicilies*, Uganda, Ukraine, Union of Soviet Socialist Republics*, Uruguay, Uzbekistan, Vanuatu, Venezuela, Vietnam, Württemberg*, Yemen, Zambia, Zimbabwe";
            return countries.ToUpper().Contains(country.ToUpper());
        }

        [ComVisible(true)]
        public String GetBranch(String keyword)
        {
            SetBranches();
            branchesArray = branchesText.Split(new String[] { "~" }, StringSplitOptions.RemoveEmptyEntries).ToList().Where(o => o.Contains(keyword.ToUpper())).ToList();
            return branchesArray.Select(o => o).FirstOrDefault();
        }
        [ComVisible(true)]
        public bool IsNumeric(string value)
        {
            double test;
            return double.TryParse(value, out test);
        }
        [ComVisible(true)]
        public bool IsAlphaNumeric(string text)
        {
            Regex objAlphaNumericPattern = new Regex("[^a-zA-Z0-9]");
            return !objAlphaNumericPattern.IsMatch(text);
        }
        [ComVisible(true)]
        public bool IsAlpha(string text)
        {
            Regex objAlphaPattern = new Regex("[^a-zA-Z]");
            return !objAlphaPattern.IsMatch(text);
        }
        [ComVisible(true)]
        public string TakeAlpha(string text)
        {
            string result = Regex.Replace(text, @"[^a-zA-Z]+", String.Empty);
            return result;

        }
        [ComVisible(true)]
        public string TakeNumeric(string text)
        {
            string result = Regex.Replace(text, @"[^0-9]+", String.Empty);
            return result;
        }

        [ComVisible(true)]
        public string RemoveSpecialChar(string str)
        {
            string[] chars = new string[] { ",", ".", "/", "!", "@", "#", "$", "%", "^", "&", "*", "'", "\"", ";", "_", "(", ")", ":", "|", "[", "]", " " };
            for (int i = 0; i < chars.Length; i++)
            {
                if (str.Contains(chars[i]))
                {
                    str = str.Replace(chars[i], "");
                }
            }
            return str;
        }


        [ComVisible(true)]
        public bool HasSpecialChar(string input)
        {
            string specialChar = @"\|!#$%&/()=?»«@£§€{}.-;'<>_,";
            foreach (var item in specialChar)
            {
                if (input.Contains(item)) return true;
            }

            return false;
        }

        [ComVisible(true)]
        public string TakePhraseWithoutSpecialChar(string text)
        {
            string[] chars = new string[] { ",", ".", "/", "!", "@", "#", "$", "%", "^", "&", "*", "'", "\"", ";", "_", "(", ")", ":", "|", "[", "]" };
            for (int i = 0; i < chars.Length; i++)
            {
                if (text.Contains(chars[i]))
                {
                    text = text.Replace(chars[i], "");
                }
            }
            return text;
        }
        [ComVisible(true)]
        public bool TestRegex(string text)
        {
            Regex objAlphaNumericPattern = new Regex("[^a-zA-Z0-9]");
            return !objAlphaNumericPattern.IsMatch(text);
        }



        //public CustomWordType[] SortWordByLeftPointer(CustomWordType[] wordsArray)
        //{
        //    return wordsArray.OrderBy(o => o.Left).ToArray();
        //}

        public void Init(long count)
        {
            customWords = new CustomWordType[count];
        }

        public void AddCustomWord(CustomWordType addword)
        {
            if (customWords == null)
                customWords = new System.Collections.Generic.List<CustomWordType>().ToArray();

            var data = customWords.ToList();
            
            data.Add(addword);
            customWords = new CustomWordType[data.Count];
            customWords = data.ToArray();
        }

        public CustomWordType[] SortWordByLeftPointer()
        {
            customWords = customWords.OrderBy(o => o.Left).ToArray();
            return customWords;
        }

        public CustomWordType GetCustomWordByIndex(long Index)
        {
            if (Index >= customWords.Count())
                return new CustomWordType();
            return customWords.OrderBy(o => o.Left).ToArray()[Index];            
        }

        void SetBranches()
        {            
            branchesText = " IMPERIAL MALL, MIRI" + "~" + "VIVACITY, KUCHING" + "~" + "BANDAR UTAMA" + "~" + "BUKIT RAJA" + "~" + "SUNWAY PYRAMID" + "~" + "KINTA CITY" + "~" + "MINES" + "~" + "KLCC" + "~" + "SEREMBAN 2" + "~" + "TEBRAU CITY" + "~" + "KEPONG" + "~" + "CHERAS SELATAN" + "~" + "BUKIT TINGGI" + "~" + "WANGSA WALK" + "~" + "MESRA MALL" + "~" + "1ST AVENUE" + "~" + "1 SHAMELIN" + "~" + "RAWANG" + "~" + "BUKIT INDAH" + "~" + "STATION 18" + "~" + "SETIAWALK" + "~" + "GURNEY PARAGON" + "~" + "CHERAS SENTRAL" + "~" + "ENCORP STRAND" + "~" + "SERI MANJUNG" + "~" + "BUKIT MERTAJAM" + "~" + "KULAI" + "~" + "JAYA" + "~" + "D'PULZE " + "~" + "AU2" + "~" + "SUNWAY PUTRA" + "~" + "TAIPING" + "~" + "KLEBANG" + "~" + "SUNWAY VELOCITY";
        }       

    }

    public class CustomWordType
    {
        public String Text { get; set; }
        public long Left { get; set; }
        public long Top { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
    }

    public class Kadc
    {
        public static void Main(string[] args)
        {
            INeilTest neilTest = new Compare();
            //string OCRText = neilTest.GetBranch("city");

            //Console.WriteLine(OCRText);

            neilTest.AddCustomWord(new CustomWordType());
            Console.Read();
        }
    }
}
