using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using HtmlAgilityPack;
using Word = Microsoft.Office.Interop.Word;

namespace WordAccess
{
    public class Show
    {
        public DateTime Date { get; set; }
        public string Organaizer { get; set; }
        public string Contacts { get; set; }

        public string  Rank { get; set; }
        public string Mono { get; set; }

        public bool IsMono
        {
            get { return !string.IsNullOrEmpty(Mono); }
        } 
    }

    class HtmlLoader
    {
        public void Load()
        {
            string calendar = "http://uku.com.ua/shows_champ/kalend_show_2017.html";

            List<Show> shows = new List<Show>();

            //HttpClient http = new HttpClient();
            //var response = http.GetByteArrayAsync(calendar);

            //String source = Encoding.GetEncoding("windows-1251").GetString(response.Result, 0, response.Result.Length - 1);
            //source = WebUtility.HtmlDecode(source);
            HtmlWeb web = new HtmlWeb()
            {
                AutoDetectEncoding = false,
                OverrideEncoding = Encoding.GetEncoding("windows-1251")
            };

            HtmlDocument resultat = web.Load(calendar);

            string organizer = string.Empty;
            string contacts = string.Empty;
            var tableRows = resultat.DocumentNode.SelectNodes("//table[@width='1140']/tr");
            for (int i = 3; i < tableRows.Count; i++)
            {
                DateTime date = DateTime.MinValue;
                Dictionary<string, bool> ranks = new Dictionary<string, bool>();

                var cells = tableRows[i].SelectNodes("td");
                foreach (var cell in cells)
                {
                    if (cell.Attributes["class"].Value == "kol_1")
                    {
                        // Date
                        if (!string.IsNullOrEmpty(cell.InnerText))
                        {
                            DateTime.TryParse(cell.InnerText, CultureInfo.CreateSpecificCulture("ru"),
                                DateTimeStyles.AssumeLocal, out date);
                        }
                    }
                    if (cell.Attributes["class"].Value == "kol_2")
                    {
                        // Organizer
                        organizer = cell.InnerText;
                    }
                    if (cell.Attributes["class"].Value == "kol_3")
                    {
                        // Description - ranks
                        if (cell.InnerText.Contains("CAC"))
                        {
                            ranks.Add("CAC", false);
                        }
                        if (cell.InnerText.Contains("Монопородні виставки:"))
                        {
                            var textNode = cell.SelectSingleNode("(p/text())[last()]");
                            if (textNode == null)
                            {
                                textNode = cell.SelectSingleNode("(text())[last()]");
                            }
                            if (textNode != null)
                            {
                                var breeds = textNode.InnerText.Trim().Split(',');
                                foreach (string breed in breeds)
                                {
                                    if (!ranks.ContainsKey(breed.Trim()))
                                    {
                                        ranks.Add(breed.Trim(), true);
                                    }
                                }
                            }
                           
                        }

                        var links = cell.SelectNodes("//a");
                        if (links != null)
                        {
                            
                        }

                        //todo: parse another cases 
                    }
                    if (cell.Attributes["class"].Value == "kol_4")
                    {
                        // Contacts
                        contacts = cell.InnerText;
                    }
                    Console.WriteLine(cell.InnerText);
                }

                foreach (KeyValuePair<string, bool> rank in ranks)
                {
                    shows.Add(new Show()
                    {
                        Date = date,
                        Organaizer = organizer,
                        Contacts = contacts,
                        Mono = rank.Value ? rank.Key : string.Empty,
                        Rank = ranks.First(r => !r.Value).Key
                    });
                }
            }
            Console.WriteLine(shows.Count);
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.GetEncoding("windows-1251");
            Console.InputEncoding = Encoding.GetEncoding("windows-1251");

            RegistrationInfo info = new RegistrationInfo()
            {
                Breed = "Rottweiler",
                Sex = "Female",
                Color = "Black&Tan",
                DOB = new DateTime(2016, 3, 12),
                Name = "Britney Bright Black",
                Pedigree = "UKU.0276614",
                Father = "Norris Lug De Burg UKU.0167434",
                Mother = "Irishka Klensmen Kliri UKU.0191057",
                Breeder = "Dmytro Chevelev",
                Owner = "Dmytro Chevelev",
                Club = "Canis (Kharkov)",
                Address = "Ukraine, 61124 Kharkov, Gagarina av. 176/4-89",
                Class = "ЮНИОРЫ",
                Mono = "Ротвейлер",
                Email = "chevelevd@gmail.com",
                Phone = "0667353399"
            };

            //SouzvivatProcessor processor = new SouzvivatProcessor("http://uku.com.ua/temporary/2017/04/30_01_harkov_cac.docx");
            //processor.Process(info);

            DergachiProcessor dProcessor = new DergachiProcessor("http://uku.com.ua/temporary/2017/05/08_dergachev.doc");
            dProcessor.Process(info);

            //fillDocument(wordApp, "http://uku.com.ua/temporary/2017/05/08_dergachev.doc", "Юниоры");
            //fillDocument(wordApp, "http://uku.com.ua/temporary/2017/05/09_dergachev.doc", "Юниоры");




        }

    }
}
