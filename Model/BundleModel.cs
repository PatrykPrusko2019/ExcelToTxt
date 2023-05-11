namespace ExcelToTxt.Model
{
    /// <summary>
    /// wiązki
    /// </summary>
    public class BundleModel : BaseModel
    {
        public override string ToString()
        {
            //FirstCylinder -> erste flasche
            return ($"{Company};{Bundle};{Internal};{GasType};{FirstCylinder};{Status};{NumerOfCylinders};{S_Owner};{Manufacturer};{ManufacturerDate};" +
                    $"{LastTesting};{FillingPress};{TestingPressure};{PeriodTrial};{Prueforgan};{Volume};{TotalWeight};{FillPressCode};{FillingWeight};{SizeCode};{Mass};{FinancialOwner}");
        }

    }
}
