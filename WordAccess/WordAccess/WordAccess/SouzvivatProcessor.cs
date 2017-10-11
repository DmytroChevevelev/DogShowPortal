using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordAccess
{
    public class SouzvivatProcessor : DocumentProcessorBase
    {
        private const string DATE = "Дата :";
        private const string PHONE = "Тел.";
        private const string EMAIL = "E-mail:";


        public SouzvivatProcessor(object fileName) : base(fileName)
        {
        }

        protected override void OnProcess(Cell cell, Cell nextCell, RegistrationInfo info)
        {
            var cellText = cell.Range.Text;
            if (cellText.Contains(DATE))
            {
                cell.Range.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}", 
                    cell.Range.Text.TrimEnd('\a').TrimEnd('\r'), 
                    DateTime.Now.ToString(DateFormat));
            }

            if (cellText.Contains(PHONE))
            {
                cell.Range.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}", 
                    cell.Range.Text.TrimEnd('\a').TrimEnd('\r') , info.Phone);
            }

            if (cellText.Contains(EMAIL))
            {
                cell.Range.Text = string.Format(CultureInfo.InvariantCulture, "{0} {1}",
                    cell.Range.Text.TrimEnd('\a').TrimEnd('\r'), info.Email);
            }

            if (cellText.Contains(info.Class) && nextCell != null)
            {
                nextCell.Range.Text = "X";
            }

            if (cellText.Contains("САС ВСЕХ ПОРОД") && nextCell != null)
            {
                nextCell.Range.Text = "X";
            }

            if (cellText.Contains(info.Mono))
            {
                cell.Range.Text = string.Format(CultureInfo.InvariantCulture, "{0}\tX",
                    cell.Range.Text.TrimEnd('\a').TrimEnd('\r'));
            }
        }
    }
}
