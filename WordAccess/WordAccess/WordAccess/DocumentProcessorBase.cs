using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace WordAccess
{
    public abstract class DocumentProcessorBase
    {
        protected const string DateFormat = "dd.MM.yyyy";

        private Word.Application wordApp;
        private Word.Document document;
        public DocumentProcessorBase(object fileName)
        {
            wordApp = new Word.Application();
            document = wordApp.Documents.Open(ref fileName);
            //document.SaveAs(string.Format(CultureInfo.InvariantCulture, @"C:\Users\Dmytro.chevelev\Downloads\{0}", document.Name));
            document.Activate();
            document.ActiveWindow.View.ReadingLayout = false;
        }

        public void Process(RegistrationInfo regInfo)
        {
            foreach (Word.Table table in document.Tables)
            {
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        try
                        {
                            var cellText = table.Cell(r, c).Range.Text;
                            Word.Cell nextCell = null;
                            try
                            {
                                nextCell = table.Cell(r, c + 1);
                            }
                            catch (Exception)
                            {
                                
                            }

                            if (nextCell != null)
                            {
                                if (cellText.Contains("Breed") || cellText.Contains("Порода"))
                                {
                                    nextCell.Range.Text = regInfo.Breed;
                                }

                                if (cellText.Contains("Sex") || cellText.Contains("Пол"))
                                {
                                    nextCell.Range.Text = regInfo.Sex;
                                }

                                if (cellText.Contains("Color") || cellText.Contains("Цвет"))
                                {
                                    nextCell.Range.Text = regInfo.Color;
                                }

                                if (cellText.Contains("birth") || cellText.Contains("рожд"))
                                {
                                    nextCell.Range.Text = regInfo.DOB.ToString(DateFormat);
                                }

                                if (cellText.Contains("of the dog") || cellText.Contains("Кличка"))
                                {
                                    nextCell.Range.Text = regInfo.Name;
                                }

                                if (cellText.Contains("Pedigree") || cellText.Contains("родос"))
                                {
                                    nextCell.Range.Text = regInfo.Pedigree;
                                }

                                if (cellText.Contains("Father") || cellText.Contains("Отец"))
                                {
                                    nextCell.Range.Text = regInfo.Father;
                                }

                                if (cellText.Contains("Mother") || cellText.Contains("Мать"))
                                {
                                    nextCell.Range.Text = regInfo.Mother;
                                }

                                if (cellText.Contains("Breeder") || cellText.Contains("Заводчик"))
                                {
                                    nextCell.Range.Text = regInfo.Breeder;
                                }

                                if (cellText.Contains("Owner") || cellText.Contains("Владелец"))
                                {
                                    nextCell.Range.Text = regInfo.Owner;
                                }

                                if (cellText.Contains("Клуб регистрации"))
                                {
                                    nextCell.Range.Text = regInfo.Club;
                                }

                                if (cellText.Contains("Address") || cellText.Contains("Адрес"))
                                {
                                    nextCell.Range.Text = regInfo.Address;
                                }

                                if (cellText.Contains(regInfo.Class))
                                {
                                    nextCell.Range.Text = "X";
                                }
                            }
                            
                            OnProcess(table.Cell(r, c), nextCell, regInfo);
                        }
                        catch (Exception)
                        {
                            //break;
                        }
                    }
                }
            }
            document.Save();
            wordApp.Visible = true;
        }

        protected abstract void OnProcess(Word.Cell cell, Word.Cell nextCell, RegistrationInfo info);
    }
}
