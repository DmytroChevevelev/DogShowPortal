using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordAccess
{
    public class DergachiProcessor : DocumentProcessorBase
    {
        private const string DATE = "Дата";
        private const string PHONE = "Телефон:";
        private const string EMAIL = "e-mail";
        public DergachiProcessor(object fileName) : base(fileName)
        {
        }

        protected override void OnProcess(Cell cell, Cell nextCell, RegistrationInfo info)
        {
            var cellText = cell.Range.Text;
            if (nextCell != null && cellText.Contains(PHONE))
            {
                nextCell.Range.Text = info.Phone;
            }

            if (nextCell != null && cellText.Contains(EMAIL))
            {
                nextCell.Range.Text = info.Email;
            }

            if (cellText.Contains(DATE))
            {
                cell.Range.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}",
                    cell.Range.Text.TrimEnd('\a').TrimEnd('\r'),
                    DateTime.Now.ToString(DateFormat));
            }
        }
    }
}
