using System.Collections.Generic;

namespace ExcelToTxt.CorrectData
{
    public class CorrectBundle : BaseCorrectData
    {
        public CorrectBundle()
        {
            Company = new List<bool>();
            Bundle = new List<bool>();
            Internal = new List<bool>();
            GasType = new List<bool>();
            FirstCylinder = new List<bool>();
            Status = new List<bool>();
            NumerOfCylinders = new List<bool>();
            S_Owner = new List<bool>();
            Manufacturer = new List<bool>();
            ManufacturerDate = new List<bool>();
            LastTesting = new List<bool>();
            FillingPress = new List<bool>();
            TestingPressure = new List<bool>();
            PeriodTrial = new List<bool>();
            Prueforgan = new List<bool>();
            Volume = new List<bool>();
            TotalWeight = new List<bool>();
            FillPressCode = new List<bool>();
            FillingWeight = new List<bool>();
            SizeCode = new List<bool>();
            Mass = new List<bool>();
            FinancialOwner = new List<bool>();
            RealWeight = new List<bool>();

            ListOfFields = new Dictionary<string, List<bool>>();

            ListOfFields.Add("Company", Company);
            ListOfFields.Add("Bundle", Bundle);
            ListOfFields.Add("Internal", Internal);
            ListOfFields.Add("GasType", GasType);
            ListOfFields.Add("FirstCylinder", FirstCylinder);
            ListOfFields.Add("Status", Status);
            ListOfFields.Add("NumerOfCylinders", NumerOfCylinders);
            ListOfFields.Add("S_Owner", S_Owner);
            ListOfFields.Add("Manufacturer", Manufacturer);
            ListOfFields.Add("ManufacturerDate", ManufacturerDate);
            ListOfFields.Add("LastTesting", LastTesting);
            ListOfFields.Add("FillingPress", FillingPress);
            ListOfFields.Add("TestingPressure", TestingPressure);
            ListOfFields.Add("PeriodTrial", PeriodTrial);
            ListOfFields.Add("Prueforgan", Prueforgan);
            ListOfFields.Add("Volume", Volume);
            ListOfFields.Add("TotalWeight", TotalWeight);
            ListOfFields.Add("FillPressCode", FillPressCode);
            ListOfFields.Add("FillingWeight", FillingWeight);
            ListOfFields.Add("SizeCode", SizeCode);
            ListOfFields.Add("Mass", Mass);
            ListOfFields.Add("FinancialOwner", FinancialOwner);
            ListOfFields.Add("RealWeight", RealWeight);
        }

    }
}
