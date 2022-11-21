using OfficeOpenXml;
using System.Text.RegularExpressions;

// start configuration by this input
string originFilePath = "C:\\Users\\MattiaCaserio\\Desktop\\associati.xlsx";
string destStreamWriter = "C:\\Users\\MattiaCaserio\\Desktop\\import_script.sql";
string sheetName = "Sorgenti";

Console.WriteLine("Run parser...");
StreamWriter sw = new StreamWriter(destStreamWriter);
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
Regex rg = new Regex("^\\d+/\\d+/\\d"); // use to check if a string in column could be a date
string startIns = "INSERT INTO [dbo].[Associates] ([Year] ,[VenueCode] ,[AssociateType] ,[Origin] ,[FirstRegistrationDate] ,[CardNo] ,[Opearation] ,[SAPCod] ,[ExportINPSDate] ,[DelegatedSignatureDate] ,[Status] ,[ResponseINPSCod] ,[AssociationRevoked] ,[INPSCod] ,[FiscalCod] ,[LastName] ,[FirstName] ,[BirhDate] ,[BirthPlace] ,[ResidencePlace] ,[Address] ,[Cap] ,[Email] ,[Phone] ,[MobilePhone] ,[Pec] ,[PIVA] ,[CompanyName] ,[LegalForm] ,[CompanyType] ,[CompanyRole] ,[AtecoCod] ,[AtecoSecondaryCod1] ,[AtecoSecondaryCod2] ,[Category] ,[AssociatesNumber] ,[Year_toDel] ,[Value_toDel] ,[Name_toDel] ,[PIVA_toDel] ,[City] ,[Address_toDel] ,[CAP_toDel] ,[Email_toDel] ,[Phone_toDel] ,[MobilePhone_toDel] ,[PEC_toDel] ,[Value1] ,[Status1] ,[Value2] ,[Status2] ,[Value3] ,[Status3] ,[Value4] ,[Status4] ,[Value5] ,[Status5] ,[PaiedInstallmentTot] ,[Federation1] ,[Federation2] ,[CategoryNationalCoordinator] ,[AtecoCod2] ,[Region] ,[AtecoCod2Desc] ,[Channel] ,[PaiedInstallment1] ,[PaiedInstallment2] ,[PaiedInstallment3] ,[PaiedInstallment4] ,[UniquePaiedInstallment] ,[Payer] ,[Prov] ,[Age] ,[AgeRange] ,[Gender]) VALUES(";
string endIns = ");";
string valuesIns = "";
using (var package = new ExcelPackage(new FileInfo(originFilePath)))
{    
    var sheet = package.Workbook.Worksheets[sheetName];
    for (var rowNum = 2; rowNum <= sheet.Dimension.End.Row; rowNum++)
    {
        var currRow = sheet.Cells[rowNum, 1, rowNum, sheet.Dimension.End.Column];
        foreach (var col in currRow)
        {
            //try to get the right type of each column's value
            bool isNull = col.Value == null;
            bool isNumber = col.Value is double;
            bool isDate = col.Value is DateTime;
            bool isString = col.Value is string;

            if (isNull) valuesIns += "null,";
            else if (isString)
            {
                //if string is a date
                if (col.Value is string s && rg.IsMatch(s)) valuesIns += $"convert(datetime, '{col.Value}', 103),";
                //if string is empty set its value to null to avoid DB errors
                else if (col.Value != null && col.Value.ToString() == "") valuesIns += "null,";
                // in this case col has a good value, nothing to do :)
                else valuesIns += $"'{col.Value?.ToString()?.Replace(@"'", @"''")}',";
            }
            else if (isDate) valuesIns += $"convert(datetime, '{col.Value}', 103),";
            else valuesIns += $"{col.Value?.ToString()?.Replace(',', '.')},";
        }
        valuesIns = valuesIns.Remove(valuesIns.Length - 1, 1);
        sw.WriteLine($"{startIns}{valuesIns}{endIns}");
        valuesIns = "";
    }
    Console.WriteLine("File Created Succesfully!");
    sw.Close();
}
