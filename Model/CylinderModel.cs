namespace ExcelToTxt.Model
{
    /// <summary>
    /// butle
    /// </summary>
    public class CylinderModel : BaseModel
    {


        public override string ToString()
        {
            return ($"{Company};{Bundle};{GasType};{Status};{S_Owner};{NumerOfCylinders};{Manufacturer};{ManufacturerNumber};{ManufacturerDate};{LastTesting};{ListType};{FillingPress};" +
                    $"{TestingPressure};{Mass};{InternalTreatment};{Material};{Volume};{NominalWeight};{RealWeight};{SizeCode};{FillPressCode};{WeightType};{Height};{Diameter};{Inlet};" +
                    $"{Destination};{Neckring};{FinancialOwner};{Connector};{Barcode};{AssetNumber};{AssetSubnumber}");
        }
    }
}
