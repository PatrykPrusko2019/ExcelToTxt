namespace ExcelToTxt.Model
{
    public abstract class BaseModel
    {
        public string Bundle { get; set; }
        public string Company { get; set; }
        public string GasType { get; set; }
        public string Status { get; set; }
        public string S_Owner { get; set; }
        public string Manufacturer { get; set; }
        public string ManufacturerNumber { get; set; }
        public string ManufacturerDate { get; set; }
        public string LastTesting { get; set; }
        public string FillingPress { get; set; }
        public string TestingPressure { get; set; }
        public string Material { get; set; }
        public string Volume { get; set; }
        public string NominalWeight { get; set; }
        public string RealWeight { get; set; }
        public string FillPressCode { get; set; }
        public string Neckring { get; set; }
        public string FinancialOwner { get; set; }

        public string Internal { get; set; }
        public string NumerOfCylinders { get; set; }
        public string FirstCylinder { get; set; }
        public string TotalWeight { get; set; }
        public string Mass { get; set; }
        public string SizeCode { get; set; }
        public string ListType { get; set; }
        public string InternalTreatment { get; set; }
        public string WeightType { get; set; }
        public string Height { get; set; }
        public string Diameter { get; set; }
        public string Inlet { get; set; }
        public string Destination { get; set; }
        public string Connector { get; set; }
        public string Barcode { get; set; }
        public string AssetNumber { get; set; }
        public string AssetSubnumber { get; set; }
        public string PeriodTrial { get; set; }
        public string Prueforgan { get; set; }
        public string FillingWeight { get; set; }
    }
}
