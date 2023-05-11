using ExcelToTxt.CorrectData;
using ExcelToTxt.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToTxt.GetCylinders
{
    public class DownloadDataCylinder
    {
        private string sheetName;

        public CorrectCylinder CorrectCylinder { get; set; }

        public DownloadDataCylinder()
        {
        }

        public DownloadDataCylinder(string sheetName)
        {
            this.sheetName = sheetName;
        }

        public String CompleteCylinderData(ISheet sheet)
        {
            string codeCountry = "";
            CylinderModel cylinder = null;
            List<CylinderModel> cylinders = new List<CylinderModel>();
            StringBuilder stringBuilder = new StringBuilder();
            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;

                var names = new Dictionary<string, int>();//pobieram dane nazw kolumn do tabeli
                IRow curROW = sheet.GetRow(0);
                for (int i = 0; i < curROW.Cells.Count; i++)
                {
                    string columnName = sheet.GetRow(0).GetCell(i).ToString();
                    if (!names.ContainsKey(columnName))
                    {
                        names.Add(columnName, i); // key -> nazwa naglowka, value -> nr komorki
                    }
                }
                names = Utils.GetActuallColumns(names); // pobiera aktualne lokalizacje kolumn 

                int correctIndex = Utils.SearchGoodIndex(sheet, curROW.Cells.Count);
                if (correctIndex == -1) { Console.WriteLine("index is not correct !!! wrong file excel !!!"); return ""; }

                for (; correctIndex < rowCount + 1; correctIndex++)
                {
                    cylinder = new CylinderModel();

                    curROW = sheet.GetRow(correctIndex);


                    int resultColumn;

                    if (names.TryGetValue("Bundle", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Bundle = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Bundle = "";
                    }

                    if (names.TryGetValue("GasType", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempStr = curROW.GetCell(resultColumn).ToString();
                        if (tempStr.ToLower().Equals("tlen")) { tempStr = "O2"; }
                        else if (tempStr.ToLower().Equals("azot")) { tempStr = "N2"; }
                        cylinder.GasType = tempStr;
                    }
                    else
                    {
                        cylinder.GasType = "";
                    }

                    if (names.TryGetValue("Status", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Status = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Status = "";
                    }

                    if (names.TryGetValue("S_Owner", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.S_Owner = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.S_Owner = "";
                    }

                    //sprawdzam czy juz puste sa pola
                    if (Utils.CheckIsNextRecordCylinder(cylinder)) break;

                    if (names.TryGetValue("Company", out resultColumn) && !(curROW.GetCell(resultColumn).ToString().Equals("")) && codeCountry == "")
                    {
                        cylinder.Company = curROW.GetCell(resultColumn).ToString();
                    } // jesli brak code country to uzupelnia po wybraniu
                    else if (codeCountry == "")
                    {
                        codeCountry = Utils.CountryCodeSelection();
                        cylinder.Company = codeCountry;
                    }
                    else
                    {
                        cylinder.Company = codeCountry;
                    }

                    if (names.TryGetValue("NumerOfCylinders", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.NumerOfCylinders = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.NumerOfCylinders = "";
                    }

                    if (names.TryGetValue("Manufacturer", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string valueTemp = curROW.GetCell(resultColumn).ToString();
                        cylinder.Manufacturer = valueTemp.Substring(0, 3); // DALMINE -> DAL
                    }
                    else
                    {
                        cylinder.Manufacturer = "";
                    }

                    if (names.TryGetValue("ManufacturerNumber", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.ManufacturerNumber = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.ManufacturerNumber = "";
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

                        cylinder.ManufacturerDate = date.ToString("dd.MM.yyyy"); // cylinder.ManufacturerDate = date.ToString("dd-MM-yyyy");

                    }
                    else
                    {
                        cylinder.ManufacturerDate = "";
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

                        cylinder.LastTesting = date.ToString("yyyy.MM.dd"); // cylinder.LastTesting = date.ToString("yyyy'/'MM'/'dd");

                    }
                    else
                    {
                        cylinder.LastTesting = "";
                    }

                    if (names.TryGetValue("FillingPress", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.FillingPress = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.FillingPress = "";
                    }

                    if (names.TryGetValue("TestingPressure", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.TestingPressure = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.TestingPressure = "";
                    }

                    if (names.TryGetValue("Material", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Material = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Material = "";
                    }

                    if (names.TryGetValue("Volume", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            cylinder.Volume = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            cylinder.Volume = tempVolume;
                        }

                    }
                    else
                    {
                        cylinder.Volume = "";
                    }

                    if (names.TryGetValue("NominalWeight", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            cylinder.NominalWeight = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            cylinder.NominalWeight = tempVolume;
                        }

                    }
                    else
                    {
                        cylinder.NominalWeight = "";
                    }

                    if (names.TryGetValue("RealWeight", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string tempVolume = curROW.GetCell(resultColumn).ToString();
                        tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);

                        if (int.TryParse(tempVolume, out int result) || double.TryParse(tempVolume.Replace('.', ','), out double result2))
                        {
                            cylinder.RealWeight = tempVolume;
                        }
                        else
                        {
                            tempVolume = curROW.GetCell(resultColumn).NumericCellValue.ToString();
                            tempVolume = Utils.GetCorrectValueDoubleToString(tempVolume);
                            cylinder.RealWeight = tempVolume;
                        }

                    }
                    else
                    {
                        cylinder.RealWeight = "";
                    }

                    if (names.TryGetValue("FillPressCode", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.FillPressCode = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.FillPressCode = "";
                    }

                    if (names.TryGetValue("Neckring", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Neckring = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Neckring = "";
                    }

                    if (names.TryGetValue("FinancialOwner", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.FinancialOwner = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.FinancialOwner = "";
                    }

                    if (names.TryGetValue("Height", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Height = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Height = "";
                    }

                    if (names.TryGetValue("Diameter", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Diameter = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Diameter = "";
                    }

                    if (names.TryGetValue("Inlet", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Inlet = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Inlet = "";
                    }

                    if (names.TryGetValue("Destination", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Destination = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Destination = "";
                    }

                    if (names.TryGetValue("Connector", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        string str = curROW.GetCell(resultColumn).ToString(); // np SMARTOP -> SMTP
                        if (str.ToLower().Contains("smartop"))
                        {
                            string newStr = str.Remove(2) + str.ElementAt(4) + str.ElementAt(6);
                            str = newStr;
                        }

                        cylinder.Connector = str;
                    }
                    else
                    {
                        cylinder.Connector = "";
                    }

                    if (names.TryGetValue("Barcode", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Barcode = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Barcode = "";
                    }

                    if (names.TryGetValue("AssetNumber", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.AssetNumber = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.AssetNumber = "";
                    }

                    if (names.TryGetValue("AssetSubnumber", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.AssetSubnumber = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.AssetSubnumber = "";
                    }

                    if (names.TryGetValue("InternalTreatment", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.InternalTreatment = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.InternalTreatment = "";
                    }

                    if (names.TryGetValue("ListType", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.ListType = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.ListType = "";
                    }

                    if (names.TryGetValue("SizeCode", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.SizeCode = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.SizeCode = "";
                    }

                    if (names.TryGetValue("WeightType", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.WeightType = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.WeightType = "";
                    }

                    if (names.TryGetValue("Mass", out resultColumn) && Utils.CheckCellIsEmpty(curROW.GetCell(resultColumn)))
                    {
                        cylinder.Mass = curROW.GetCell(resultColumn).ToString();
                    }
                    else
                    {
                        cylinder.Mass = "";
                    }



                    cylinders.Add(cylinder);

                    // tak ma zwracac -> 5;3160855;AR;V;PL;DAL;12887278;01-03-2021;2021/03/01;300;450;ST;50.0;69.3;69.3;3;W80;ALPL

                    //stringBuilder.Append(cylinder.ToString());
                    // stringBuilder.AppendLine();

                }

                CorrectCylinder = new CorrectCylinder();
                stringBuilder.Append(Utils.GetCorrectCylinders(cylinders, CorrectCylinder, names));
            }

            return stringBuilder.ToString();
        }


        public string CheckFieldsAreComplete()
        {
            return CorrectCylinder.CheckLists();
        }
    }
}
