using ExcelToTxt.CorrectData;
using ExcelToTxt.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace ExcelToTxt
{
    public class Utils
    {

        public static Dictionary<string, int> GetActuallColumns(Dictionary<string, int> names)
        {
            var actuallColumnsList = new Dictionary<string, int>();
            foreach (var item in names)
            {
                string name = "";
                if (item.Key.Equals("Właściciel") || item.Key.Equals("Własciciel") || item.Key.Equals("Wlaściciel") || item.Key.Equals("Wlasciciel"))
                {
                    name = item.Key; // only Upper letter first
                }
                else
                {
                    name = item.Key.ToLower();
                }

                if (name.Equals("nr wiazki") || name.Equals("bundle") || name.Equals("bundle number") || name.Equals("bundle no.") || name.Equals("bündelnummer") || name.Equals("bundelnummer") || name.Equals("buendelnummer"))
                {
                    if (!actuallColumnsList.ContainsKey("Bundle"))
                    {
                        actuallColumnsList.Add("Bundle", item.Value);
                    }

                }
                else if (name.Equals("numer wewnetrzny") || name.Equals("nr wewnetrzny") || name.Equals("internal number") || name.Equals("internal") || name.Equals("internal nummber") || name.Equals("aga_nummer") ||
                         name.Equals("aga nummer") || name.Equals("aga_nummber") || name.Equals("aga nummber"))
                {
                    if (!actuallColumnsList.ContainsKey("Internal"))
                    {
                        actuallColumnsList.Add("Internal", item.Value);
                    }

                }
                else if (name.Equals("company") || name.Equals("firma"))
                {
                    if (!actuallColumnsList.ContainsKey("Company"))
                    {
                        actuallColumnsList.Add("Company", item.Value);
                    }

                }
                else if (name.Equals("gas type") || name.Equals("gaseart") || name.Equals("typ gazu"))
                {
                    if (!actuallColumnsList.ContainsKey("GasType"))
                    {
                        actuallColumnsList.Add("GasType", item.Value);
                    }

                }
                else if (name.Equals("status") || name.Equals("status:") || name.Equals("status "))
                {
                    if (!actuallColumnsList.ContainsKey("Status"))
                    {
                        actuallColumnsList.Add("Status", item.Value);
                    }

                }
                else if (name.Equals("s_owner") || name.Equals("user company") || name.Equals("verfuger") || name.Equals("right to use") || name.Equals("right to usel") || name.Equals("właściciel") || name.Equals("wlaściciel") || name.Equals("wlasciciel") || name.Equals("własciciel"))
                {
                    if (!actuallColumnsList.ContainsKey("S_Owner"))
                    {
                        actuallColumnsList.Add("S_Owner", item.Value);
                    }

                }
                else if (name.Equals("manufacturer number") || name.Equals("manufact.number") || name.Equals("erzeugernummer") || name.Equals("erzeugernr.") || name.Equals("numer producenta"))
                {
                    if (!actuallColumnsList.ContainsKey("ManufacturerNumber"))
                    {
                        actuallColumnsList.Add("ManufacturerNumber", item.Value);
                    }
                }
                else if (name.Equals("manufacturer") || name.Equals("erzeuger") || name.Equals("manufacturer:") || name.Equals("producent"))
                {
                    if (!actuallColumnsList.ContainsKey("Manufacturer"))
                    {
                        actuallColumnsList.Add("Manufacturer", item.Value);
                    }

                }
                else if (name.Equals("manufacturer date") || name.Equals("manufact. date") || name.Equals("erzeugungsdatum") || name.Equals("d_erzeugung") || name.Equals("manufactured data") || name.Equals("manufactured date") || name.Equals("data producji"))
                {
                    if (!actuallColumnsList.ContainsKey("ManufacturerDate"))
                    {
                        actuallColumnsList.Add("ManufacturerDate", item.Value);
                    }

                }
                else if (name.Equals("last testing") || name.Equals("last tesiting") || name.Equals("letzte prufung") || name.Equals("letzte prüfung") || name.Equals("d_l_pruefung") || name.Equals("letze prufung") || name.Equals("letze prfung") || name.Equals("last testing ") || name.Equals("data legalizacji")
                        || name.Equals("last testing data"))
                {
                    if (!actuallColumnsList.ContainsKey("LastTesting"))
                    {
                        actuallColumnsList.Add("LastTesting", item.Value);
                    }

                }
                else if (name.Equals("filling pressure") || name.Equals("filling_pressure") || name.Equals("filling press") || name.Equals("filling press.") || name.Equals("fuelldruck") || name.Equals("füelldruck") || name.Equals("fulldruck") || name.Equals("fülldruck") || name.Equals("cisnienie rob.") || name.Equals("ciśnienie rob."))
                {
                    if (!actuallColumnsList.ContainsKey("FillingPress"))
                    {
                        actuallColumnsList.Add("FillingPress", item.Value);
                    }

                }
                else if (name.Equals("testing pressure") || name.Equals("testing_pressure") || name.Equals("pruefdruck") || name.Equals("prüefdruck") || name.Equals("prufdruck") || name.Equals("prüfdruck") || name.Equals("ciśnienie próbne") || name.Equals("ciśnienie probne") || name.Equals("cisnienie próbne") || name.Equals("cisnienie probne"))
                {
                    if (!actuallColumnsList.ContainsKey("TestingPressure"))
                    {
                        actuallColumnsList.Add("TestingPressure", item.Value);
                    }

                }
                else if (name.Equals("material") || name.Equals("flaschen material") || name.Equals("rodz. mat") || name.Equals("rodzaj mat"))
                {
                    if (!actuallColumnsList.ContainsKey("Material"))
                    {
                        actuallColumnsList.Add("Material", item.Value);
                    }

                }
                else if (name.Equals("volume") || name.Equals("inhalt") || name.Equals("flaschen volumen") || name.Equals("flaschen-volumen") || name.Equals("pojemność") || name.Equals("pojemnosc") || name.Equals("pojemnośc") || name.Equals("pojemnosć"))
                {
                    if (!actuallColumnsList.ContainsKey("Volume"))
                    {
                        actuallColumnsList.Add("Volume", item.Value);
                    }

                }
                else if (name.Equals("nominal weight") || name.Equals("sollgewicht") || name.Equals("brutto weight"))
                {
                    if (!actuallColumnsList.ContainsKey("NominalWeight"))
                    {
                        actuallColumnsList.Add("NominalWeight", item.Value);
                    }

                }
                else if (name.Equals("real weight") || name.Equals("istgewicht") || name.Equals("ist-gewicht") || name.Equals("tara weight") || name.Equals("waga butli") || name.Equals("gesamtgewicht"))
                {
                    if (!actuallColumnsList.ContainsKey("RealWeight"))
                    {
                        actuallColumnsList.Add("RealWeight", item.Value);
                    }

                }
                else if (name.Equals("fillpressCode") || name.Equals("fill.press.code") || name.Equals("fulldruck (code)") || name.Equals("fülldruck (code)") ||
                         name.Equals("fülldruck code") || name.Equals("fulldruck code") || name.Equals("filling pressure code") || name.Equals("druck") || name.Equals("drück") || name.Equals("working pressure") || name.Equals("kod do pełnienia") || name.Equals("kod do pelnienia") || name.Equals("kod do pelnienia ") || name.Equals("kod do pełnienia "))
                {
                    if (!actuallColumnsList.ContainsKey("FillPressCode"))
                    {
                        actuallColumnsList.Add("FillPressCode", item.Value);
                    }

                }
                else if (name.Equals("neckring") || name.Equals("neck ring thread") || name.Equals("neckring thread") || name.Equals("halsringgwinde") || name.Equals("halsring gwinde") || name.Equals("halsring-gewinde") ||
                         name.Equals("hr_gewindetyp") || name.Equals("halsring-gewind") || name.Equals("'halsring-gewind") || name.Equals("halsring gewind") || name.Equals("'halsring gewind"))
                {
                    if (!actuallColumnsList.ContainsKey("Neckring"))
                    {
                        actuallColumnsList.Add("Neckring", item.Value);

                    }
                }
                else if (name.Equals("financial owner") || name.Equals("fin_owner") || name.Equals("fin owner") || name.Equals("eigentummer") || name.Equals("eigentümmer") || name.Equals("eigentumer") || name.Equals("eigentümer") || name.Equals("financial owner:") || name.Equals("Właściciel") || name.Equals("Wlaściciel") || name.Equals("Wlasciciel") || name.Equals("Własciciel"))
                {
                    if (!actuallColumnsList.ContainsKey("FinancialOwner"))
                    {
                        actuallColumnsList.Add("FinancialOwner", item.Value);

                    }

                }
                else if (name.Equals("numer of cylinders") || name.Equals("number of cylinders") || name.Equals("cylinder number") || name.Equals("cylinder no.") || name.Equals("flaschennummer") || name.Equals("flaschenanzahl") || name.Equals("no of cylinders ") || name.Equals("No of cylinders"))
                {
                    if (!actuallColumnsList.ContainsKey("NumerOfCylinders"))
                    {
                        actuallColumnsList.Add("NumerOfCylinders", item.Value);
                    }

                }
                else if (name.Equals("1st cylinder") || name.Equals("1-cylinder number") || name.Equals("erste_flasche") || name.Equals("erste flasche") || name.Equals("1 st cylinder ") || name.Equals("1 st cylinder") || name.Equals("1 cylinder number") || name.Equals("1 -cylinder number"))
                {
                    if (!actuallColumnsList.ContainsKey("FirstCylinder"))
                    {
                        actuallColumnsList.Add("FirstCylinder", item.Value);
                    }

                }
                else if (name.Equals("total weight") || name.Equals("gesamtgewicht") || name.Equals("lączny ciężar"))
                {
                    if (!actuallColumnsList.ContainsKey("TotalWeight"))
                    {
                        actuallColumnsList.Add("TotalWeight", item.Value);
                    }

                }
                else if (name.Equals("mass") || name.Equals("masse"))
                {
                    if (!actuallColumnsList.ContainsKey("Mass"))
                    {
                        actuallColumnsList.Add("Mass", item.Value);
                    }

                }
                else if (name.Equals("size code") || name.Equals("inhalt code") || name.Equals("code_inh") || name.Equals("code inh") ||
                         name.Equals("inhalt (code)") || name.Equals("flaschen-größe") || name.Equals("flaschen-größe cod") ||
                         name.Equals("flaschen-große cod") || name.Equals("flaschen größe cod") || name.Equals("flaschen große cod") ||
                         name.Equals("flaschen-große") || name.Equals("flaschen-größe (cod)") || name.Equals("flaschen-große (cod)") || name.Equals("kod do pojemności ")
                         || name.Equals("kod do pojemnosci ") || name.Equals("kod do pojemności") || name.Equals("kod do pojemnosci"))
                {
                    if (!actuallColumnsList.ContainsKey("SizeCode"))
                    {
                        actuallColumnsList.Add("SizeCode", item.Value);
                    }

                }
                else if (name.Equals("list type") || name.Equals("listen typ"))
                {
                    if (!actuallColumnsList.ContainsKey("ListType"))
                    {
                        actuallColumnsList.Add("ListType", item.Value);
                    }

                }
                else if (name.Equals("internal treatm.") || name.Equals("internal treatment") || name.Equals("innenbehandlung"))
                {
                    if (!actuallColumnsList.ContainsKey("InternalTreatment"))
                    {
                        actuallColumnsList.Add("InternalTreatment", item.Value);
                    }

                }
                else if (name.Equals("weight type") || name.Equals("gewichtstyp") || name.Equals("gew_typ") || name.Equals("gew typ"))
                {
                    if (!actuallColumnsList.ContainsKey("WeightType"))
                    {
                        actuallColumnsList.Add("WeightType", item.Value);
                    }

                }
                else if (name.Equals("height") || name.Equals("höhe") || name.Equals("höehe") || name.Equals("hohe") || name.Equals("hoehe"))
                {
                    if (!actuallColumnsList.ContainsKey("Height"))
                    {
                        actuallColumnsList.Add("Height", item.Value);
                    }

                }
                else if (name.Equals("diameter") || name.Equals("durchmesser") || name.Equals("flaschen-durchmesser") || name.Equals("flaschen durchmesser"))
                {
                    if (!actuallColumnsList.ContainsKey("Diameter"))
                    {
                        actuallColumnsList.Add("Diameter", item.Value);
                    }

                }
                else if (name.Equals("inlet") || name.Equals("inlet thread") || name.Equals("'einschaub-gewind") || name.Equals("'einschaub gewind") ||
                         name.Equals("einschaub gewind") || name.Equals("einschaub-gewind") || name.Equals("gewindetyp") || name.Equals("gewinde typ") ||
                         name.Equals("einschaub-gewinde") || name.Equals("einschaub gewinde"))
                {
                    if (!actuallColumnsList.ContainsKey("Inlet"))
                    {
                        actuallColumnsList.Add("Inlet", item.Value);
                    }

                }
                else if (name.Equals("destination") || name.Equals("herkunftsland") || name.Equals("country of origin"))
                {
                    if (!actuallColumnsList.ContainsKey("Destination"))
                    {
                        actuallColumnsList.Add("Destination", item.Value);
                    }

                }
                else if (name.Equals("connector type") || name.Equals("connector") || name.Equals("behart_subtyp") || name.Equals("behart subtyp"))
                {
                    if (!actuallColumnsList.ContainsKey("Connector"))
                    {
                        actuallColumnsList.Add("Connector", item.Value);
                    }

                }
                else if (name.Equals("barcode") || name.Equals("strichcode") || name.Equals("nowy barcode"))
                {
                    if (!actuallColumnsList.ContainsKey("Barcode"))
                    {
                        actuallColumnsList.Add("Barcode", item.Value);
                    }

                }
                else if (name.Equals("asset number") || name.Equals("asset_number") || name.Equals("Inventarnummer") || name.Equals("Ressourcennummer"))
                {
                    if (!actuallColumnsList.ContainsKey("AssetNumber"))
                    {
                        actuallColumnsList.Add("AssetNumber", item.Value);
                    }

                }
                else if (name.Equals("asset_subnumber") || name.Equals("asset subnumber"))
                {
                    if (!actuallColumnsList.ContainsKey("AssetSubnumber"))
                    {
                        actuallColumnsList.Add("AssetSubnumber", item.Value);
                    }

                }
                else if (name.Equals("pruefperiode") || name.Equals("period trial") || name.Equals("maint.period"))
                {
                    if (!actuallColumnsList.ContainsKey("PeriodTrial"))
                    {
                        actuallColumnsList.Add("PeriodTrial", item.Value);
                    }

                }
                else if (name.Equals("prueforgan"))
                {
                    if (!actuallColumnsList.ContainsKey("Prueforgan"))
                    {
                        actuallColumnsList.Add("Prueforgan", item.Value);
                    }

                }
                else if (name.Equals("filling weight") || name.Equals("fuellgewicht"))
                {
                    if (!actuallColumnsList.ContainsKey("FillingWeight"))
                    {
                        actuallColumnsList.Add("FillingWeight", item.Value);
                    }

                }


            }
            return actuallColumnsList;
        }

        public static string GetCorrectValueDoubleToString(string tempVolume)
        {
            tempVolume = tempVolume.Contains(",") ? tempVolume.Replace(',', '.') : tempVolume;
            tempVolume = tempVolume.Contains(".") ? tempVolume : tempVolume + ".0";
            return tempVolume;
        }

        public static int SearchGoodIndex(ISheet sheet, int count)
        {
            int wrongIndex = -1;

            for (int i = 1; i < sheet.LastRowNum; i++)
            {
                IRow curROW = sheet.GetRow(i);
                for (int j = 0; j < count; j++)
                {
                    if (CheckCellIsEmpty(curROW.GetCell(j)))
                    {
                        string valueCell = curROW.GetCell(j).ToString();
                        if (int.TryParse(valueCell, out int result))
                        {
                            return i; // search number
                        }

                    }
                }


            }

            return wrongIndex;
        }

        /// <summary>
        /// uzupelnia dany rekord
        /// </summary>
        /// <param name="cylinders"></param>
        /// <param name="correctCylinder"></param>
        /// <returns></returns>
        public static string GetCorrectCylinders(List<CylinderModel> cylinders, CorrectCylinder correctCylinder, Dictionary<string, int> names)
        {
            StringBuilder stringBuilder = new StringBuilder();

            foreach (var cylinder in cylinders)
            {
                stringBuilder.Append(IfValueEmptyCheckCylinder(cylinder, correctCylinder, names));
                stringBuilder.AppendLine();
            }

            return stringBuilder.ToString();
        }



        /// <summary>
        /// checks if the fields are empty, if not empty, fill in the field
        /// if 1 -> true
        /// if -1 -> false
        /// </summary>
        /// <param name="cylinder"></param>
        /// <returns></returns>
        private static string IfValueEmptyCheckCylinder(CylinderModel cylinder, CorrectCylinder correctCylinder, Dictionary<string, int> names)
        {

            StringBuilder stringBuilder = new StringBuilder();

            if (!String.IsNullOrEmpty(cylinder.Company)) { stringBuilder.Append(cylinder.Company + ';'); correctCylinder.Company.Add(true); } else { correctCylinder.Company.Add(false); }

            foreach (var item in names)
            {
                switch (item.Key)
                {
                    case "Bundle":
                        if (!String.IsNullOrEmpty(cylinder.Bundle)) { stringBuilder.Append(cylinder.Bundle + ';'); correctCylinder.Bundle.Add(true); } else { correctCylinder.Bundle.Add(false); }
                        break;

                    case "GasType":
                        if (!String.IsNullOrEmpty(cylinder.GasType)) { stringBuilder.Append(cylinder.GasType + ';'); correctCylinder.GasType.Add(true); } else { correctCylinder.GasType.Add(false); }
                        break;

                    case "Status":
                        if (!String.IsNullOrEmpty(cylinder.Status)) { stringBuilder.Append(cylinder.Status + ';'); correctCylinder.Status.Add(true); } else { correctCylinder.Status.Add(false); }
                        break;

                    case "S_Owner":
                        if (!String.IsNullOrEmpty(cylinder.S_Owner)) { stringBuilder.Append(cylinder.S_Owner + ';'); correctCylinder.S_Owner.Add(true); } else { correctCylinder.S_Owner.Add(false); }
                        break;

                    case "NumerOfCylinders":
                        if (!String.IsNullOrEmpty(cylinder.NumerOfCylinders)) { stringBuilder.Append(cylinder.NumerOfCylinders + ';'); correctCylinder.NumerOfCylinders.Add(true); } else { correctCylinder.NumerOfCylinders.Add(false); }
                        break;

                    case "Manufacturer":
                        if (!String.IsNullOrEmpty(cylinder.Manufacturer)) { stringBuilder.Append(cylinder.Manufacturer + ';'); correctCylinder.Manufacturer.Add(true); } else { correctCylinder.Manufacturer.Add(false); }
                        break;

                    case "ManufacturerNumber":
                        if (!String.IsNullOrEmpty(cylinder.ManufacturerNumber)) { stringBuilder.Append(cylinder.ManufacturerNumber + ';'); correctCylinder.ManufacturerNumber.Add(true); } else { correctCylinder.ManufacturerNumber.Add(false); }
                        break;

                    case "ManufacturerDate":
                        if (!String.IsNullOrEmpty(cylinder.ManufacturerDate)) { stringBuilder.Append(cylinder.ManufacturerDate + ';'); correctCylinder.ManufacturerDate.Add(true); } else { correctCylinder.ManufacturerDate.Add(false); }
                        break;

                    case "LastTesting":
                        if (!String.IsNullOrEmpty(cylinder.LastTesting)) { stringBuilder.Append(cylinder.LastTesting + ';'); correctCylinder.LastTesting.Add(true); } else { correctCylinder.LastTesting.Add(false); }
                        break;

                    case "ListType":
                        if (!String.IsNullOrEmpty(cylinder.ListType)) { stringBuilder.Append(cylinder.ListType + ';'); correctCylinder.ListType.Add(true); } else { correctCylinder.ListType.Add(false); }
                        break;

                    case "FillingPress":
                        if (!String.IsNullOrEmpty(cylinder.FillingPress)) { stringBuilder.Append(cylinder.FillingPress + ';'); correctCylinder.FillingPress.Add(true); } else { correctCylinder.FillingPress.Add(false); }
                        break;

                    case "TestingPressure":
                        if (!String.IsNullOrEmpty(cylinder.TestingPressure)) { stringBuilder.Append(cylinder.TestingPressure + ';'); correctCylinder.TestingPressure.Add(true); } else { correctCylinder.TestingPressure.Add(false); }
                        break;

                    case "Mass":
                        if (!String.IsNullOrEmpty(cylinder.Mass)) { stringBuilder.Append(cylinder.Mass + ';'); correctCylinder.Mass.Add(true); } else { correctCylinder.Mass.Add(false); }
                        break;

                    case "InternalTreatment":
                        if (!String.IsNullOrEmpty(cylinder.InternalTreatment)) { stringBuilder.Append(cylinder.InternalTreatment + ';'); correctCylinder.InternalTreatment.Add(true); } else { correctCylinder.InternalTreatment.Add(false); }
                        break;

                    case "Material":
                        if (!String.IsNullOrEmpty(cylinder.Material)) { stringBuilder.Append(cylinder.Material + ';'); correctCylinder.Material.Add(true); } else { correctCylinder.Material.Add(false); }
                        break;

                    case "Volume":
                        if (!String.IsNullOrEmpty(cylinder.Volume)) { stringBuilder.Append(cylinder.Volume + ';'); correctCylinder.Volume.Add(true); } else { correctCylinder.Volume.Add(false); }
                        break;

                    case "NominalWeight":
                        if (!String.IsNullOrEmpty(cylinder.NominalWeight)) { stringBuilder.Append(cylinder.NominalWeight + ';'); correctCylinder.NominalWeight.Add(true); } else { correctCylinder.NominalWeight.Add(false); }
                        break;

                    case "RealWeight":
                        if (!String.IsNullOrEmpty(cylinder.RealWeight)) { stringBuilder.Append(cylinder.RealWeight + ';'); correctCylinder.RealWeight.Add(true); } else { correctCylinder.RealWeight.Add(false); }
                        break;

                    case "SizeCode":
                        if (!String.IsNullOrEmpty(cylinder.SizeCode)) { stringBuilder.Append(cylinder.SizeCode + ';'); correctCylinder.SizeCode.Add(true); } else { correctCylinder.SizeCode.Add(false); }
                        break;

                    case "FillPressCode":
                        if (!String.IsNullOrEmpty(cylinder.FillPressCode)) { stringBuilder.Append(cylinder.FillPressCode + ';'); correctCylinder.FillPressCode.Add(true); } else { correctCylinder.FillPressCode.Add(false); }
                        break;

                    case "WeightType":
                        if (!String.IsNullOrEmpty(cylinder.WeightType)) { stringBuilder.Append(cylinder.WeightType + ';'); correctCylinder.WeightType.Add(true); } else { correctCylinder.WeightType.Add(false); }
                        break;

                    case "Height":
                        if (!String.IsNullOrEmpty(cylinder.Height)) { stringBuilder.Append(cylinder.Height + ';'); correctCylinder.Height.Add(true); } else { correctCylinder.Height.Add(false); }
                        break;

                    case "Diameter":
                        if (!String.IsNullOrEmpty(cylinder.Diameter)) { stringBuilder.Append(cylinder.Diameter + ';'); correctCylinder.Diameter.Add(true); } else { correctCylinder.Diameter.Add(false); }
                        break;

                    case "Inlet":
                        if (!String.IsNullOrEmpty(cylinder.Inlet)) { stringBuilder.Append(cylinder.Inlet + ';'); correctCylinder.Inlet.Add(true); } else { correctCylinder.Inlet.Add(false); }
                        break;

                    case "Destination":
                        if (!String.IsNullOrEmpty(cylinder.Destination)) { stringBuilder.Append(cylinder.Destination + ';'); correctCylinder.Destination.Add(true); } else { correctCylinder.Destination.Add(false); }
                        break;

                    case "Neckring":
                        if (!String.IsNullOrEmpty(cylinder.Neckring)) { stringBuilder.Append(cylinder.Neckring + ';'); correctCylinder.Neckring.Add(true); } else { correctCylinder.Neckring.Add(false); }
                        break;

                    case "FinancialOwner":
                        if (!String.IsNullOrEmpty(cylinder.FinancialOwner)) { stringBuilder.Append(cylinder.FinancialOwner + ';'); correctCylinder.FinancialOwner.Add(true); } else { correctCylinder.FinancialOwner.Add(false); }
                        break;

                    case "Connector":
                        if (!String.IsNullOrEmpty(cylinder.Connector)) { stringBuilder.Append(cylinder.Connector + ';'); correctCylinder.Connector.Add(true); } else { correctCylinder.Connector.Add(false); }
                        break;

                    case "Barcode":
                        if (!String.IsNullOrEmpty(cylinder.Barcode)) { stringBuilder.Append(cylinder.Barcode + ';'); correctCylinder.Barcode.Add(true); } else { correctCylinder.Barcode.Add(false); }
                        break;

                    case "AssetNumber":
                        if (!String.IsNullOrEmpty(cylinder.AssetNumber)) { stringBuilder.Append(cylinder.AssetNumber + ';'); correctCylinder.AssetNumber.Add(true); } else { correctCylinder.AssetNumber.Add(false); }
                        break;

                    case "AssetSubnumber":
                        if (!String.IsNullOrEmpty(cylinder.AssetSubnumber)) { stringBuilder.Append(cylinder.AssetSubnumber + ';'); correctCylinder.AssetSubnumber.Add(true); } else { correctCylinder.AssetSubnumber.Add(false); }
                        break;



                }
            }


            //if (!String.IsNullOrEmpty(cylinder.Company)) { stringBuilder.Append(cylinder.Company + ';'); correctCylinder.Company.Add(true); } else { correctCylinder.Company.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Bundle)) { stringBuilder.Append(cylinder.Bundle + ';'); correctCylinder.Bundle.Add(true); } else { correctCylinder.Bundle.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.GasType)) { stringBuilder.Append(cylinder.GasType + ';'); correctCylinder.GasType.Add(true); } else { correctCylinder.GasType.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Status)) { stringBuilder.Append(cylinder.Status + ';'); correctCylinder.Status.Add(true); } else { correctCylinder.Status.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.S_Owner)) { stringBuilder.Append(cylinder.S_Owner + ';'); correctCylinder.S_Owner.Add(true); } else { correctCylinder.S_Owner.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.NumerOfCylinders)) { stringBuilder.Append(cylinder.NumerOfCylinders + ';'); correctCylinder.NumerOfCylinders.Add(true); } else { correctCylinder.NumerOfCylinders.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Manufacturer)) { stringBuilder.Append(cylinder.Manufacturer + ';'); correctCylinder.Manufacturer.Add(true); } else { correctCylinder.Manufacturer.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.ManufacturerNumber)) { stringBuilder.Append(cylinder.ManufacturerNumber + ';'); correctCylinder.ManufacturerNumber.Add(true); } else { correctCylinder.ManufacturerNumber.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.ManufacturerDate)) { stringBuilder.Append(cylinder.ManufacturerDate + ';'); correctCylinder.ManufacturerDate.Add(true); } else { correctCylinder.ManufacturerDate.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.LastTesting)) { stringBuilder.Append(cylinder.LastTesting + ';'); correctCylinder.LastTesting.Add(true); } else { correctCylinder.LastTesting.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.ListType)) { stringBuilder.Append(cylinder.ListType + ';'); correctCylinder.ListType.Add(true); } else { correctCylinder.ListType.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.FillingPress)) { stringBuilder.Append(cylinder.FillingPress + ';'); correctCylinder.FillingPress.Add(true); } else { correctCylinder.FillingPress.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.TestingPressure)) { stringBuilder.Append(cylinder.TestingPressure + ';'); correctCylinder.TestingPressure.Add(true); } else { correctCylinder.TestingPressure.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Mass)) { stringBuilder.Append(cylinder.Mass + ';'); correctCylinder.Mass.Add(true); } else { correctCylinder.Mass.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.InternalTreatment)) { stringBuilder.Append(cylinder.InternalTreatment + ';'); correctCylinder.InternalTreatment.Add(true); } else { correctCylinder.InternalTreatment.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Material)) { stringBuilder.Append(cylinder.Material + ';'); correctCylinder.Material.Add(true); } else { correctCylinder.Material.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Volume)) { stringBuilder.Append(cylinder.Volume + ';'); correctCylinder.Volume.Add(true); } else { correctCylinder.Volume.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.NominalWeight)) { stringBuilder.Append(cylinder.NominalWeight + ';'); correctCylinder.NominalWeight.Add(true); } else { correctCylinder.NominalWeight.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.RealWeight)) { stringBuilder.Append(cylinder.RealWeight + ';'); correctCylinder.RealWeight.Add(true); } else { correctCylinder.RealWeight.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.SizeCode)) { stringBuilder.Append(cylinder.SizeCode + ';'); correctCylinder.SizeCode.Add(true); } else { correctCylinder.SizeCode.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.FillPressCode)) { stringBuilder.Append(cylinder.FillPressCode + ';'); correctCylinder.FillPressCode.Add(true); } else { correctCylinder.FillPressCode.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.WeightType)) { stringBuilder.Append(cylinder.WeightType + ';'); correctCylinder.WeightType.Add(true); } else { correctCylinder.WeightType.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Height)) { stringBuilder.Append(cylinder.Height + ';'); correctCylinder.Height.Add(true); } else { correctCylinder.Height.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Diameter)) { stringBuilder.Append(cylinder.Diameter + ';'); correctCylinder.Diameter.Add(true); } else { correctCylinder.Diameter.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Inlet)) { stringBuilder.Append(cylinder.Inlet + ';'); correctCylinder.Inlet.Add(true); } else { correctCylinder.Inlet.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Destination)) { stringBuilder.Append(cylinder.Destination + ';'); correctCylinder.Destination.Add(true); } else { correctCylinder.Destination.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Neckring)) { stringBuilder.Append(cylinder.Neckring + ';'); correctCylinder.Neckring.Add(true); } else { correctCylinder.Neckring.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.FinancialOwner)) { stringBuilder.Append(cylinder.FinancialOwner + ';'); correctCylinder.FinancialOwner.Add(true); } else { correctCylinder.FinancialOwner.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Connector)) { stringBuilder.Append(cylinder.Connector + ';'); correctCylinder.Connector.Add(true); } else { correctCylinder.Connector.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.Barcode)) { stringBuilder.Append(cylinder.Barcode + ';'); correctCylinder.Barcode.Add(true); } else { correctCylinder.Barcode.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.AssetNumber)) { stringBuilder.Append(cylinder.AssetNumber + ';'); correctCylinder.AssetNumber.Add(true); } else { correctCylinder.AssetNumber.Add(false); }
            //if (!String.IsNullOrEmpty(cylinder.AssetSubnumber)) { stringBuilder.Append(cylinder.AssetSubnumber + ';'); correctCylinder.AssetSubnumber.Add(true); } else { correctCylinder.AssetSubnumber.Add(false); }



            stringBuilder.Remove(stringBuilder.Length - 1, 1);
            return stringBuilder.ToString();
        }

        public static string[] GetCorrectData(string valuesDate)
        {
            string[] strings = valuesDate.Split("-");

            if (strings.Length == 1)
            {
                strings = null;
                strings = valuesDate.Split("/");
            }
            if (strings.Length == 1)
            {
                strings = null;
                strings = valuesDate.Split(@"\");
            }
            if (strings.Length == 1)
            {
                strings = null;
                strings = valuesDate.Split(".");
            }

            return strings;
        }



        public static string GetCorrectBundles(List<BundleModel> bundles, CorrectBundle correctBundle, Dictionary<string, int> names)
        {
            StringBuilder stringBuilder = new StringBuilder();

            foreach (var bundle in bundles)
            {
                stringBuilder.Append(IfValueEmptyCheckBundle(bundle, correctBundle, names));
                stringBuilder.AppendLine();
            }

            return stringBuilder.ToString();
        }



        /// <summary>
        /// checks if the fields are empty, if not empty, fill in the field
        /// </summary>
        /// <param name="bundle"></param>
        /// <param name="correctBundle"></param>
        /// <returns></returns>
        private static string IfValueEmptyCheckBundle(BundleModel bundle, CorrectBundle correctBundle, Dictionary<string, int> names)
        {
            StringBuilder stringBuilder = new StringBuilder();

            if (!String.IsNullOrEmpty(bundle.Company)) { stringBuilder.Append(bundle.Company + ';'); correctBundle.Company.Add(true); } else { correctBundle.Company.Add(false); }

            foreach (var item in names)
            {
                switch (item.Key)
                {

                    case "Bundle":
                        if (!String.IsNullOrEmpty(bundle.Bundle)) { stringBuilder.Append(bundle.Bundle + ';'); correctBundle.Bundle.Add(true); } else { correctBundle.Bundle.Add(false); }
                        break;

                    case "Internal":
                        if (!String.IsNullOrEmpty(bundle.Internal)) { stringBuilder.Append(bundle.Internal + ';'); correctBundle.Internal.Add(true); } else { correctBundle.Internal.Add(false); }
                        break;

                    case "GasType":
                        if (!String.IsNullOrEmpty(bundle.GasType)) { stringBuilder.Append(bundle.GasType + ';'); correctBundle.GasType.Add(true); } else { correctBundle.GasType.Add(false); }
                        break;

                    case "FirstCylinder":
                        if (!String.IsNullOrEmpty(bundle.FirstCylinder)) { stringBuilder.Append(bundle.FirstCylinder + ';'); correctBundle.FirstCylinder.Add(true); } else { correctBundle.FirstCylinder.Add(false); }
                        break;

                    case "Status":
                        if (!String.IsNullOrEmpty(bundle.Status)) { stringBuilder.Append(bundle.Status + ';'); correctBundle.Status.Add(true); } else { correctBundle.Status.Add(false); }
                        break;

                    case "NumerOfCylinders":
                        if (!String.IsNullOrEmpty(bundle.NumerOfCylinders)) { stringBuilder.Append(bundle.NumerOfCylinders + ';'); correctBundle.NumerOfCylinders.Add(true); } else { correctBundle.NumerOfCylinders.Add(false); }
                        break;

                    case "S_Owner":
                        if (!String.IsNullOrEmpty(bundle.S_Owner)) { stringBuilder.Append(bundle.S_Owner + ';'); correctBundle.S_Owner.Add(true); } else { correctBundle.S_Owner.Add(false); }
                        break;

                    case "Manufacturer":
                        if (!String.IsNullOrEmpty(bundle.Manufacturer)) { stringBuilder.Append(bundle.Manufacturer + ';'); correctBundle.Manufacturer.Add(true); } else { correctBundle.Manufacturer.Add(false); }
                        break;

                    case "ManufacturerDate":
                        if (!String.IsNullOrEmpty(bundle.ManufacturerDate)) { stringBuilder.Append(bundle.ManufacturerDate + ';'); correctBundle.ManufacturerDate.Add(true); } else { correctBundle.ManufacturerDate.Add(false); }
                        break;

                    case "LastTesting":
                        if (!String.IsNullOrEmpty(bundle.LastTesting)) { stringBuilder.Append(bundle.LastTesting + ';'); correctBundle.LastTesting.Add(true); } else { correctBundle.LastTesting.Add(false); }
                        break;

                    case "FillingPress":
                        if (!String.IsNullOrEmpty(bundle.FillingPress)) { stringBuilder.Append(bundle.FillingPress + ';'); correctBundle.FillingPress.Add(true); } else { correctBundle.FillingPress.Add(false); }
                        break;

                    case "TestingPressure":
                        if (!String.IsNullOrEmpty(bundle.TestingPressure)) { stringBuilder.Append(bundle.TestingPressure + ';'); correctBundle.TestingPressure.Add(true); } else { correctBundle.TestingPressure.Add(false); }
                        break;

                    case "PeriodTrial":
                        if (!String.IsNullOrEmpty(bundle.PeriodTrial)) { stringBuilder.Append(bundle.PeriodTrial + ';'); correctBundle.PeriodTrial.Add(true); } else { correctBundle.PeriodTrial.Add(false); }
                        break;

                    case "Prueforgan":
                        if (!String.IsNullOrEmpty(bundle.Prueforgan)) { stringBuilder.Append(bundle.Prueforgan + ';'); correctBundle.Prueforgan.Add(true); } else { correctBundle.Prueforgan.Add(false); }
                        break;

                    case "Volume":
                        if (!String.IsNullOrEmpty(bundle.Volume)) { stringBuilder.Append(bundle.Volume + ';'); correctBundle.Volume.Add(true); } else { correctBundle.Volume.Add(false); }
                        break;

                    case "TotalWeight":
                        if (!String.IsNullOrEmpty(bundle.TotalWeight)) { stringBuilder.Append(bundle.TotalWeight + ';'); correctBundle.TotalWeight.Add(true); } else { correctBundle.TotalWeight.Add(false); }
                        break;

                    case "FillPressCode":
                        if (!String.IsNullOrEmpty(bundle.FillPressCode)) { stringBuilder.Append(bundle.FillPressCode + ';'); correctBundle.FillPressCode.Add(true); } else { correctBundle.FillPressCode.Add(false); }
                        break;

                    case "FillingWeight":
                        if (!String.IsNullOrEmpty(bundle.FillingWeight)) { stringBuilder.Append(bundle.FillingWeight + ';'); correctBundle.FillingWeight.Add(true); } else { correctBundle.FillingWeight.Add(false); }
                        break;

                    case "SizeCode":
                        if (!String.IsNullOrEmpty(bundle.SizeCode)) { stringBuilder.Append(bundle.SizeCode + ';'); correctBundle.SizeCode.Add(true); } else { correctBundle.SizeCode.Add(false); }
                        break;

                    case "Mass":
                        if (!String.IsNullOrEmpty(bundle.Mass)) { stringBuilder.Append(bundle.Mass + ';'); correctBundle.Mass.Add(true); } else { correctBundle.Mass.Add(false); }
                        break;

                    case "FinancialOwner":
                        if (!String.IsNullOrEmpty(bundle.FinancialOwner)) { stringBuilder.Append(bundle.FinancialOwner + ';'); correctBundle.FinancialOwner.Add(true); } else { correctBundle.FinancialOwner.Add(false); }
                        break;

                    case "RealWeight":
                        if (!String.IsNullOrEmpty(bundle.RealWeight)) { stringBuilder.Append(bundle.RealWeight + ';'); correctBundle.RealWeight.Add(true); } else { correctBundle.RealWeight.Add(false); }
                        break;

                }
            }

            //if (!String.IsNullOrEmpty(bundle.Company)) { stringBuilder.Append(bundle.Company + ';'); correctBundle.Company.Add(true); } else { correctBundle.Company.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Bundle)) { stringBuilder.Append(bundle.Bundle + ';'); correctBundle.Bundle.Add(true); } else { correctBundle.Bundle.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Internal)) { stringBuilder.Append(bundle.Internal + ';'); correctBundle.Internal.Add(true); } else { correctBundle.Internal.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.GasType)) { stringBuilder.Append(bundle.GasType + ';'); correctBundle.GasType.Add(true); } else { correctBundle.GasType.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.FirstCylinder)) { stringBuilder.Append(bundle.FirstCylinder + ';'); correctBundle.FirstCylinder.Add(true); } else { correctBundle.FirstCylinder.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Status)) { stringBuilder.Append(bundle.Status + ';'); correctBundle.Status.Add(true); } else { correctBundle.Status.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.NumerOfCylinders)) { stringBuilder.Append(bundle.NumerOfCylinders + ';'); correctBundle.NumerOfCylinders.Add(true); } else { correctBundle.NumerOfCylinders.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.S_Owner)) { stringBuilder.Append(bundle.S_Owner + ';'); correctBundle.S_Owner.Add(true); } else { correctBundle.S_Owner.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Manufacturer)) { stringBuilder.Append(bundle.Manufacturer + ';'); correctBundle.Manufacturer.Add(true); } else { correctBundle.Manufacturer.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.ManufacturerDate)) { stringBuilder.Append(bundle.ManufacturerDate + ';'); correctBundle.ManufacturerDate.Add(true); } else { correctBundle.ManufacturerDate.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.LastTesting)) { stringBuilder.Append(bundle.LastTesting + ';'); correctBundle.LastTesting.Add(true); } else { correctBundle.LastTesting.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.FillingPress)) { stringBuilder.Append(bundle.FillingPress + ';'); correctBundle.FillingPress.Add(true); } else { correctBundle.FillingPress.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.TestingPressure)) { stringBuilder.Append(bundle.TestingPressure + ';'); correctBundle.TestingPressure.Add(true); } else { correctBundle.TestingPressure.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.PeriodTrial)) { stringBuilder.Append(bundle.PeriodTrial + ';'); correctBundle.PeriodTrial.Add(true); } else { correctBundle.PeriodTrial.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Prueforgan)) { stringBuilder.Append(bundle.Prueforgan + ';'); correctBundle.Prueforgan.Add(true); } else { correctBundle.Prueforgan.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Volume)) { stringBuilder.Append(bundle.Volume + ';'); correctBundle.Volume.Add(true); } else { correctBundle.Volume.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.TotalWeight)) { stringBuilder.Append(bundle.TotalWeight + ';'); correctBundle.TotalWeight.Add(true); } else { correctBundle.TotalWeight.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.FillPressCode)) { stringBuilder.Append(bundle.FillPressCode + ';'); correctBundle.FillPressCode.Add(true); } else { correctBundle.FillPressCode.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.FillingWeight)) { stringBuilder.Append(bundle.FillingWeight + ';'); correctBundle.FillingWeight.Add(true); } else { correctBundle.FillingWeight.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.SizeCode)) { stringBuilder.Append(bundle.SizeCode + ';'); correctBundle.SizeCode.Add(true); } else { correctBundle.SizeCode.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.Mass)) { stringBuilder.Append(bundle.Mass + ';'); correctBundle.Mass.Add(true); } else { correctBundle.Mass.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.FinancialOwner)) { stringBuilder.Append(bundle.FinancialOwner + ';'); correctBundle.FinancialOwner.Add(true); } else { correctBundle.FinancialOwner.Add(false); }
            //if (!String.IsNullOrEmpty(bundle.RealWeight)) { stringBuilder.Append(bundle.RealWeight + ';'); correctBundle.RealWeight.Add(true); } else { correctBundle.RealWeight.Add(false); }


            stringBuilder.Remove(stringBuilder.Length - 1, 1);
            return stringBuilder.ToString();
        }

        public static DateTime GetCorrectDate(string[] strings, int count, string columnName)
        {
            int day = -1;
            int month = -1;
            int year = -1; ;

            switch (count)
            {
                case 2:

                    month = Utils.CheckMonth(strings);
                    if (month != -1)
                    {
                        if (strings[0].Length != 4 && strings[1].Length != 4)
                        {
                            foreach (var item in strings)
                            {
                                if (int.Parse(item) != month) day = int.Parse(item);
                            }
                        }
                    }
                    else
                    {
                        day = 1;
                        foreach (var item in strings)
                        {
                            if (item.Length == 4) year = int.Parse(item);
                            else if (month == -1) month = int.Parse(item);
                        }
                    }

                    break;

                case 3:

                    month = Utils.CheckMonth(strings);

                    for (int i = 0; i < strings.Length; i++)
                    {
                        if (i == 1 && strings[i].Length == 2 && month == -1) month = int.Parse(strings[i]);
                        else if (strings[i].Length == 4) year = int.Parse(strings[i]);
                        else if (strings[i].Length == 2) day = int.Parse(strings[i]);
                    }


                    if (month > 12 || month < 1)
                    {
                        int temp = day;
                        day = month;
                        month = temp;
                    }
                    break;
            }

            DateTime date;
            if (!(year.ToString().Length == 4 && month > 0 && month < 13 && day > 0 && day < 32))
            {
                MessageBox.Show($"Column name: {columnName}, wrong format Date -> year or month or day !!! -> exit application");
                Environment.Exit(1);
                date = new DateTime();
            }
            else
            {
                date = new DateTime(year, month, day);
            }

            return date;
        }

        private static int CheckMonth(string[] strings)
        {
            for (int i = 0; i < strings.Length; i++)
            {
                string monthStr = strings[i].ToUpper();
                if (Months.STY.ToString().Equals(monthStr)) return (int)Months.STY;
                else if (Months.LUT.ToString().Equals(monthStr)) return (int)Months.LUT;
                else if (Months.MAR.ToString().Equals(monthStr)) return (int)Months.MAR;
                else if (Months.KWI.ToString().Equals(monthStr)) return (int)Months.KWI;
                else if (Months.MAJ.ToString().Equals(monthStr)) return (int)Months.MAJ;
                else if (Months.CZE.ToString().Equals(monthStr)) return (int)Months.CZE;
                else if (Months.LIP.ToString().Equals(monthStr)) return (int)Months.LIP;
                else if (Months.SIE.ToString().Equals(monthStr)) return (int)Months.SIE;
                else if (Months.WRZ.ToString().Equals(monthStr)) return (int)Months.WRZ;
                else if (Months.PAZ.ToString().Equals(monthStr) || "PAŹ".ToString().Equals(monthStr)) return (int)Months.PAZ;
                else if (Months.LIS.ToString().Equals(monthStr)) return (int)Months.LIS;
                else if (Months.GRU.ToString().Equals(monthStr)) return (int)Months.GRU;
            }
            return -1;
        }

        public static string CountryCodeSelection()
        {
            string resultCode = "";

            WindowCodeCountry windowCodeCountry = new WindowCodeCountry();
            windowCodeCountry.ShowDialog();

            resultCode = windowCodeCountry.CountryCode.Code.ToString();

            return resultCode;
        }

        public static bool CheckCellIsEmpty(ICell cell)
        {

            try
            {
                if (string.IsNullOrEmpty(cell.ToString()))
                {
                    return false;
                }
                else if (cell.ToString().Equals("--") || cell.ToString().Equals("-"))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        public static bool CheckIsNextRecordCylinder(CylinderModel cylinder)
        {
            if (String.IsNullOrEmpty(cylinder.Bundle) && String.IsNullOrEmpty(cylinder.S_Owner) && String.IsNullOrEmpty(cylinder.GasType) &&
                String.IsNullOrEmpty(cylinder.Status))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool CheckIsNextRecordBundle(BundleModel bundle)
        {
            if (String.IsNullOrEmpty(bundle.Bundle) && String.IsNullOrEmpty(bundle.Internal) && String.IsNullOrEmpty(bundle.GasType) &&
                 String.IsNullOrEmpty(bundle.ManufacturerNumber) && String.IsNullOrEmpty(bundle.Status))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

}
