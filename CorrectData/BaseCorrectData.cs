using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToTxt.CorrectData
{
    public abstract class BaseCorrectData
    {
        public List<bool> Company { get; set; }
        public List<bool> Bundle { get; set; }
        public List<bool> GasType { get; set; }
        public List<bool> Status { get; set; }
        public List<bool> S_Owner { get; set; }
        public List<bool> NumerOfCylinders { get; set; }
        public List<bool> Manufacturer { get; set; }
        public List<bool> ManufacturerNumber { get; set; }
        public List<bool> ManufacturerDate { get; set; }
        public List<bool> LastTesting { get; set; }
        public List<bool> ListType { get; set; }
        public List<bool> FillingPress { get; set; }
        public List<bool> TestingPressure { get; set; }
        public List<bool> Mass { get; set; }
        public List<bool> InternalTreatment { get; set; }
        public List<bool> Material { get; set; }
        public List<bool> Volume { get; set; }
        public List<bool> NominalWeight { get; set; }
        public List<bool> RealWeight { get; set; }
        public List<bool> SizeCode { get; set; }
        public List<bool> FillPressCode { get; set; }
        public List<bool> WeightType { get; set; }
        public List<bool> Height { get; set; }
        public List<bool> Diameter { get; set; }
        public List<bool> Inlet { get; set; }
        public List<bool> Destination { get; set; }
        public List<bool> Neckring { get; set; }
        public List<bool> FinancialOwner { get; set; }
        public List<bool> Connector { get; set; }
        public List<bool> Barcode { get; set; }
        public List<bool> AssetNumber { get; set; }
        public List<bool> AssetSubnumber { get; set; }
        public List<bool> Internal { get; set; }
        public List<bool> FirstCylinder { get; set; }
        public List<bool> TotalWeight { get; set; }
        public List<bool> FillingWeight { get; set; }
        public List<bool> PeriodTrial { get; set; }
        public List<bool> Prueforgan { get; set; }

        public Dictionary<string, List<bool>> ListOfFields { get; set; }

        public string CheckLists()
        {
            int count = 0;
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append("");

            foreach (var item in ListOfFields)
            {
                count = item.Value.Where(x => x == false).Count();
                string emptyField = CheckEmptyValue(item.Key, count);
                if (emptyField != "")
                {
                    stringBuilder.Append(emptyField);
                    stringBuilder.AppendLine();
                }
            }

            return stringBuilder.ToString();
        }

        private string CheckEmptyValue(string noCompleteField, int count)
        {
            if (count != Company.Count && count != 0) return $"{noCompleteField} -> {count} no values";
            else return "";
        }
    }
}
