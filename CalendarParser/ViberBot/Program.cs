using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ViberBot
{
    class Program
    {
        static void Main(string[] args)
        {
            Task.Run(() => {
                var bot = new ViberBot();
                while (true)
                {
                    bot.Run();
                }
            }).Wait();
        }
    }
}
