using System.Collections.Generic;

namespace ExcelToTxt.CorrectData
{
    public class CorrectCylinder : BaseCorrectData
    {
        public CorrectCylinder()
        {
            Company = new List<bool>();
            Bundle = new List<bool>();
            GasType = new List<bool>();
            Status = new List<bool>();
            S_Owner = new List<bool>();
            NumerOfCylinders = new List<bool>();
            Manufacturer = new List<bool>();
            ManufacturerNumber = new List<bool>();
            ManufacturerDate = new List<bool>();
            LastTesting = new List<bool>();
            ListType = new List<bool>();
            FillingPress = new List<bool>();
            TestingPressure = new List<bool>();
            Mass = new List<bool>();
            InternalTreatment = new List<bool>();
            Material = new List<bool>();
            Volume = new List<bool>();
            NominalWeight = new List<bool>();
            RealWeight = new List<bool>();
            SizeCode = new List<bool>();
            FillPressCode = new List<bool>();
            WeightType = new List<bool>();
            Height = new List<bool>();
            Diameter = new List<bool>();
            Inlet = new List<bool>();
            Destination = new List<bool>();
            Neckring = new List<bool>();
            FinancialOwner = new List<bool>();
            Connector = new List<bool>();
            Barcode = new List<bool>();
            AssetNumber = new List<bool>();
            AssetSubnumber = new List<bool>();
            ListOfFields = new Dictionary<string, List<bool>>();

            ListOfFields.Add("Company", Company);
            ListOfFields.Add("Bundle", Bundle);
            ListOfFields.Add("GasType", GasType);
            ListOfFields.Add("Status", Status);
            ListOfFields.Add("S_Owner", S_Owner);
            ListOfFields.Add("NumerOfCylinders", NumerOfCylinders);
            ListOfFields.Add("Manufacturer", Manufacturer);
            ListOfFields.Add("ManufacturerNumber", ManufacturerNumber);
            ListOfFields.Add("ManufacturerDate", ManufacturerDate);
            ListOfFields.Add("LastTesting", LastTesting);
            ListOfFields.Add("ListType", ListType);
            ListOfFields.Add("FillingPress", FillingPress);
            ListOfFields.Add("TestingPressure", TestingPressure);
            ListOfFields.Add("Mass", Mass);
            ListOfFields.Add("InternalTreatment", InternalTreatment);
            ListOfFields.Add("Material", Material);
            ListOfFields.Add("Volume", Volume);
            ListOfFields.Add("NominalWeight", NominalWeight);
            ListOfFields.Add("RealWeight", RealWeight);
            ListOfFields.Add("SizeCode", SizeCode);
            ListOfFields.Add("FillPressCode", FillPressCode);
            ListOfFields.Add("WeightType", WeightType);
            ListOfFields.Add("Height", Height);
            ListOfFields.Add("Diameter", Diameter);
            ListOfFields.Add("Inlet", Inlet);
            ListOfFields.Add("Destination", Destination);
            ListOfFields.Add("Neckring", Neckring);
            ListOfFields.Add("FinancialOwner", FinancialOwner);
            ListOfFields.Add("Connector", Connector);
            ListOfFields.Add("Barcode", Barcode);
            ListOfFields.Add("AssetNumber", AssetNumber);
            ListOfFields.Add("AssetSubnumber", AssetSubnumber);


        }


    }
}
