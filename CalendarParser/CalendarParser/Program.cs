using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using HtmlAgilityPack;

namespace CalendarParser
{
    [Serializable]
    public class Organization
    {
        public Organization()
        {
            Shows = new List<Show>();
        }

        public string Name { get; set; }
        public Contact Contact { get; set; }
        public List<Show> Shows { get; internal set; }
    }

    public class Link
    {
        public string Text { get; set; }
        public string Address { get; set; }
    }

    public class Contact
    {
        public Contact()
        {
            Phones = new List<string>();
            Links = new List<Link>();
        }
        public string City { get; set; }
        public List<string> Phones { get; internal set; }
        public List<Link> Links { get; internal set; }
    }

    [Serializable]
    public class Show
    {
        public DateTime Date { get; set; }
        public string Description { get; set; }
        public string Rank { get; set; }
        public string Mono { get; set; }

        public bool IsMono
        {
            get { return !string.IsNullOrEmpty(Mono); }
        }
    }

    class HtmlLoader
    {
        public List<Organization> CalendarList { get; private set; }

        public HtmlLoader()
        {
            this.CalendarList = new List<Organization>();
        }

        public void LoadByOrganization()
        {
            
            string calendar = //"http://uku.com.ua/shows_champ/kalend_show_2018.html";
            "http://uku.com.ua/shows_champ/kalend_show_2017.html";
            HtmlWeb web = new HtmlWeb()
            {
                AutoDetectEncoding = false,
                OverrideEncoding = Encoding.GetEncoding("windows-1251")
            };

            HtmlDocument document = web.Load(calendar);

            var organizations = document.DocumentNode.SelectNodes("//table[@width='1140']/tr/td[@class='kol_2']");
            foreach (var organization in organizations)
            {
                var nodes = organization.SelectNodes("b");
                string organizationName = nodes.Aggregate(string.Empty, (current, node) => current + (HttpUtility.HtmlDecode(node.InnerText)));

                int iRows = 1;
                if (organization.Attributes.Contains("rowspan"))
                {
                    var rows = organization.Attributes["rowspan"].Value;
                    int.TryParse(rows, NumberStyles.Integer, CultureInfo.InvariantCulture, out iRows);
                }

                Organization org = new Organization()
                {
                    Name = organizationName
                };
                
                //get contact information
                var contactNode = organization.ParentNode.SelectSingleNode("td[@class='kol_4']");
                org.Contact = GetContactInformation(contactNode);
                

                //get dates
                var dateNode = organization.ParentNode.SelectSingleNode("td[@class='kol_1']");
                DateTime date = GetDate(dateNode);

                var descriptionNode = organization.ParentNode.SelectSingleNode("td[@class='kol_3']");
                string description = getDescription(descriptionNode);
                org.Shows.Add(new Show() { Date = date, Description = description });

                for (int r = 1; r < iRows; r++)
                {
                    dateNode = dateNode.ParentNode.SelectSingleNode("following-sibling::tr/td[@class='kol_1']");
                    descriptionNode =
                        descriptionNode.ParentNode.SelectSingleNode("following-sibling::tr/td[@class='kol_3']");
                    date = GetDate(dateNode);
                    description = getDescription(descriptionNode);
                    if (date != DateTime.MinValue)
                    {
                        org.Shows.Add(new Show() { Date = date, Description = description });
                    }
                }

                Console.WriteLine(organizationName);
                CalendarList.Add(org);
            }

        }

        private string getDescription(HtmlNode node)
        {
            if (node != null && !string.IsNullOrEmpty(node.InnerText))
            {
                return HttpUtility.HtmlDecode(node.InnerText);
            }
            return string.Empty;
        }

        private Contact GetContactInformation(HtmlNode contactNode)
        {
            Contact c = new Contact();
            if (contactNode != null)
            {
                var city = contactNode.SelectSingleNode("b[0]");
                if (city != null && !string.IsNullOrEmpty(city.InnerText))
                {
                    c.City = city.InnerText;
                }

                var phones = contactNode.SelectNodes("text()");
                if (phones != null)
                {
                    foreach (HtmlNode phone in phones)
                    {
                        string p = HttpUtility.HtmlDecode(phone.InnerText);
                        p = p.Trim(new []{'\n', ' ', ','});
                        if (!string.IsNullOrEmpty(p))
                        {
                            c.Phones.Add(p);
                        }
                        
                    }
                }

                var links = contactNode.SelectNodes("a");
                if (links != null)
                {
                    foreach (HtmlNode link in links)
                    {
                        c.Links.Add(new Link()
                        {
                            Address = link.Attributes["href"].Value,
                            Text = link.InnerText
                        });
                    }
                }
            }
            
            return c;
        }

        private DateTime GetDate(HtmlNode node)
        {
            DateTime date = DateTime.MinValue;
            if (node != null && !string.IsNullOrEmpty(node.InnerText))
            {
                DateTime.TryParse(node.InnerText, CultureInfo.CreateSpecificCulture("ru"),
                    DateTimeStyles.AssumeLocal, out date);
            }
            return date;
        }

        public void Load()
        {
            string calendar = "http://uku.com.ua/shows_champ/kalend_show_2017.html";

            List<Show> shows = new List<Show>();

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

            HtmlLoader loader = new HtmlLoader();
            loader.LoadByOrganization();
        }
    }
}
