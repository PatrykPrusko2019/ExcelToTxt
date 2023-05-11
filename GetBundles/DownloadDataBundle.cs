using ExcelToTxt.CorrectData;
using ExcelToTxt.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToTxt.GetBundles
{
    public class DownloadDataBundle
    {
        private string Name;

        public CorrectBundle CorrectBundle { get; set; }

        public DownloadDataBundle()
        {
        }

        public DownloadDataBundle(string name)
        {
            Name = name;
        }

        public string CompleteBundleData(ISheet sheet)
        {
            string codeCountry = "";
            BundleModel bundle = null;
            List<BundleModel> bundles = new List<BundleModel>();
            StringBuilder stringBuilder = new StringBuilder();

            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;

                var names = new Dictionary<string, int>();//pobieram dane nazw kolumn do tabeli
                IRow curROW = sheet.GetRow(0);
                for (int i = 0; i < curROW.Cells.Count; i++)
                {
                    string columnName = sheet.GetRow(0).GetCell(i).StringCellValue;
                    if (!names.ContainsKey(columnName))
                        names.Add(columnName, i); // key -> nazwa naglowka, value -> nr komorki
                }

                names = Utils.GetActuallColumns(names); // pobiera aktualne lokalizacje kolumn 

                int correctIndex = Utils.SearchGoodIndex(sheet, curROW.Cells.Count);
                if (correctIndex == -1) { Console.WriteLine("index is not correct !!! wrong file excel !!!"); return ""; }

                for (; correctIndex < rowCount + 1; correctIndex++)
                {
                    bundle = new BundleModel();

                    curROW = sheet.GetRow(correctIndex);

                    int resultColumn;

                    if (names.TryGetValue("Bundle", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.Bundle = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.Bundle = "";
                    }

                    if (names.TryGetValue("Internal", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.Internal = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.Internal = "";
                    }

                    if (names.TryGetValue("GasType", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempStr = curROW.GetCell(resultColumn).ToString();
                        if (tempStr.ToLower().Equals("tlen")) { tempStr = "O2"; }
                        else if (tempStr.ToLower().Equals("azot")) { tempStr = "N2"; }
                        bundle.GasType = tempStr;

                    }
                    else
                    {
                        bundle.GasType = "";
                    }

                    if (names.TryGetValue("Manufacturer", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.Manufacturer = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.Manufacturer = "";
                    }

                    if (names.TryGetValue("ManufacturerNumber", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.ManufacturerNumber = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.ManufacturerNumber = "";
                    }

                    if (names.TryGetValue("Status", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.Status = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.Status = "";
                    }

                    //sprawdzam czy sa juz puste pola
                    if (Utils.CheckIsNextRecordBundle(bundle)) break;

                    if (names.TryGetValue("Company", out resultColumn) && !(curROW.GetCell(resultColumn).ToString().Equals("")) && codeCountry == "")
                    {
                        bundle.Company = curROW.GetCell(resultColumn).ToString();
                    } // jesli brak code country to uzupelnia po wybraniu
                    else if (codeCountry == "")
                    {
                        bundle.Company = Utils.CountryCodeSelection();
                        codeCountry = bundle.Company;
                    }
                    else
                    {
                        bundle.Company = codeCountry;
                    }

                    if (names.TryGetValue("S_Owner", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.S_Owner = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.S_Owner = "";
                    }

                    if (names.TryGetValue("NumerOfCylinders", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.NumerOfCylinders = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.NumerOfCylinders = "";
                    }


                    if (names.TryGetValue("FirstCylinder", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        bundle.FirstCylinder = tempVolume.Contains("-") ? tempVolume.Replace("-", "") : tempVolume;
                    }
                    else
                    {
                        bundle.FirstCylinder = "";
                    }

                    if (names.TryGetValue("ManufacturerDate", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {

                        string valuesDate = curROW.GetCell(resultColumn).ToString();
                        string[] strings = Utils.GetCorrectData(valuesDate);

                        DateTime date = new DateTime();
                        if (strings.Length == 3)
                        {
                            date = Utils.GetCorrectDate(strings, strings.Length, "ManufacturerDate");

                        }
                        else if (strings.Length == 2)
                        {
                            date = Utils.GetCorrectDate(strings, strings.Length, "ManufacturerDate");
                        }

                        bundle.ManufacturerDate = date.ToString("dd.MM.yyyy"); // bundle.ManufacturerDate = date.ToString("dd-MM-yyyy");

                    }
                    else
                    {
                        bundle.ManufacturerDate = "";
                    }

                    if (names.TryGetValue("LastTesting", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {

                        string valuesDate = curROW.GetCell(resultColumn).ToString();
                        string[] strings = Utils.GetCorrectData(valuesDate);
                        DateTime date = new DateTime();
                        if (strings.Length == 3)
                        {
                            date = Utils.GetCorrectDate(strings, strings.Length, "LastTesting");
                        }
                        else if (strings.Length == 2)
                        {
                            date = Utils.GetCorrectDate(strings, strings.Length, "LastTesting");
                        }

                        bundle.LastTesting = date.ToString("yyyy.MM.dd"); // bundle.LastTesting = date.ToString("yyyy'/'MM'/'dd");

                    }
                    else
                    {
                        bundle.LastTesting = "";
                    }

                    if (names.TryGetValue("FillPressCode", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.FillPressCode = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.FillPressCode = "";
                    }


                    if (names.TryGetValue("FillingPress", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        int index = tempVolume.IndexOf("/");
                        bundle.FillingPress = tempVolume.Contains("/") ? tempVolume.Remove(index) : tempVolume;
                    }
                    else
                    {
                        bundle.FillingPress = "";
                    }

                    if (names.TryGetValue("TestingPressure", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.TestingPressure = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.TestingPressure = "";
                    }

                    if (names.TryGetValue("Volume", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            bundle.Volume = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            bundle.Volume = tempVolume;
                        }
                    }
                    else
                    {
                        bundle.Volume = "";
                    }

                    if (names.TryGetValue("NominalWeight", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            bundle.NominalWeight = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            bundle.NominalWeight = tempVolume;
                        }
                    }
                    else
                    {
                        bundle.NominalWeight = "";
                    }

                    if (names.TryGetValue("TotalWeight", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            bundle.TotalWeight = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            bundle.TotalWeight = tempVolume;
                        }
                    }
                    else
                    {
                        bundle.TotalWeight = "";
                    }


                    if (names.TryGetValue("RealWeight", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            bundle.RealWeight = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            bundle.RealWeight = tempVolume;
                        }
                    }
                    else
                    {
                        bundle.RealWeight = "";
                    }


                    if (names.TryGetValue("FinancialOwner", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.FinancialOwner = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.FinancialOwner = "";
                    }

                    if (names.TryGetValue("PeriodTrial", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.PeriodTrial = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.PeriodTrial = "";
                    }

                    if (names.TryGetValue("Prueforgan", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.Prueforgan = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.Prueforgan = "";
                    }

                    if (names.TryGetValue("FillingWeight", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.FillingWeight = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.FillingWeight = "";
                    }

                    if (names.TryGetValue("SizeCode", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.SizeCode = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.SizeCode = "";
                    }

                    if (names.TryGetValue("Mass", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        bundle.Mass = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        bundle.Mass = "";
                    }


                    bundles.Add(bundle);

                    // tak ma zwracac -> 5;3160855;AR;V;PL;DAL;12887278;01-03-2021;2021/03/01;300;450;ST;50.0;69.3;69.3;3;W80;ALPL
                    // Console.WriteLine(bundle);
                }

                CorrectBundle = new CorrectBundle();
                stringBuilder.Append(Utils.GetCorrectBundles(bundles, CorrectBundle, names));
            }

            return stringBuilder.ToString();
        }


        public string CheckFieldsAreComplete()
        {
            return CorrectBundle.CheckLists();
        }
    }
}
