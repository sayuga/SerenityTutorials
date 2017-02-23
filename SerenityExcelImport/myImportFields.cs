namespace myProject.myImportFields
{
    using Serenity.Data;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Default.Entities;
    using System.Data;
    using System.IO;
    using myImpHelp = myImportHelper.ExcelImportHelper;
    

    public class ExcelImportHelper
    {


        /// <summary>
        /// Enum managing the myImportEntry field type
        /// </summary>
        public enum entryType
        {
            String = 1,
            Decimal = 2,
            Bool = 3,
            Int = 4,
            dateTime = 5,
            matHertz = 6,
            matPhase = 7,
            matFurnished = 8,
            matInstalled = 9,
            oneValJoin = 10,
            multiValJoin = 11
        }

        /// <summary>
        /// Casts the imported value into the proper type. 
        /// </summary>
        /// <param name="myVal">Value being imported</param>
        /// <param name="myErrors">Error Log</param>
        /// <param name="expectedHeader">Header Name</param>
        /// <param name="row">Current Row</param>
        /// <param name="myField">Row Field </param>
        /// <param name="myType">Data Type for Value</param>
        /// <returns>Casted Value</returns>        
        public static object myImportEntry(List<object> myVal, List<string> myErrors, List<string> expectedHeader, int row, entryType myType, IDbConnection myConnection)
        {
            object myField = null;
            object val = myVal[0];
            try
            {
                switch (myType)
                {
                    case entryType.String:
                        myField = Convert.ToString(val);
                        break;
                    case entryType.Decimal:
                        myField = Convert.ToDecimal(val);
                        break;
                    case entryType.Bool:
                        myField = Convert.ToBoolean(val);
                        break;
                    case entryType.Int:
                        myField = Convert.ToInt32(val);
                        break;

                    case entryType.dateTime:
                        myField = Convert.ToDateTime(val);
                        break;
                    case entryType.oneValJoin:
                        myField = oneValJoin(myConnection, myErrors, row, expectedHeader[0], val);
                        break;
                    case entryType.multiValJoin:
                        myField = multiValJoin(myConnection, myErrors, row, expectedHeader, myVal);
                        break;

                    default:
                        myField = null;
                        break;
                }
                return myField;
            }
            catch (Exception e)
            {
                foreach (string exh in expectedHeader)
                {
                    myErrors.Add(exh + " on row " + row + " had a Exception : " + e.Message);
                }

                return myField;
            }
        }

        /// <summary>
        /// Import Joint Field with multiple Required fields
        /// </summary>
        /// <param name="myConnection">DB Connection</param>
        /// <param name="myErrors">Error Log</param>
        /// <param name="row">Current Row</param>
        /// <param name="expectedHeader">List of System Headers being referenced</param>
        /// <param name="val">List of Values being Merged</param>        
        /// <returns>ID Field Value of existing or new entry</returns>
        public static Int32 multiValJoin(IDbConnection myConnection, List<String> myErrors, int row, List<string>expectedHeader, List<dynamic> val)
        {
            string[] wsField = new string[] { Convert.ToString(val[0]), Convert.ToString(val[1]) };            
            var g = AddressLogRow.Fields;
            var currRow = myConnection.TryFirst<AddressLogRow>(q => q.Select(g.CustomerId).Where(g.Floor == wsField[0] & g.Room== wsField[1])); //checks if field exists

            if (currRow == null)
            {
                try
                {
                    var newRow = myConnection.InsertAndGetID(new AddressLogRow { Floor = wsField[0], Room = wsField[1] }); //create new entry                                
                    myErrors.Add("Warning: On Row " + row + ": Entity with name " + wsField[0] + " and sub-name "+ wsField[1]+ "was not found so it was added.");
                    return Convert.ToInt32(newRow);
                }
                catch (IOException e)
                {
                    myErrors.Add("IOException for Fields " + expectedHeader[0]+" and "+ expectedHeader[1] + "  on row " + row + " : " + e.Message);                    
                }
            }
            return currRow.CustomerId.Value;
        }

        /// <summary>
        /// Import Joint Field with Single Required field
        /// </summary>
        /// <param name="myConnection">DB Connection</param>
        /// <param name="myErrors">Error Log</param>
        /// <param name="row">Current Row</param>
        /// <param name="expectedHeader">Expected System Header of Value</param>
        /// <param name="val">Required Field </param>
        /// <returns></returns>
        public static Int32 oneValJoin(IDbConnection myConnection, List<String> myErrors, int row, string expectedHeader, dynamic val)
        {
            string wsField = Convert.ToString(val);            
            var g = AddressLogRow.Fields;
            var currRow = myConnection.TryFirst<AddressLogRow>(q => q.Select(g.AddressLogId).Where(g.Floor == wsField));

            if (currRow == null)
            {
                try
                {
                    var newRow = myConnection.InsertAndGetID(new AddressLogRow { Floor = wsField, Room  = "(Imported Field)" }); //create new entry                                
                    myErrors.Add("Warning: On Row " + row + ": Entity with name " + wsField+ " was not found so it was added.");
                    return Convert.ToInt32(newRow);
                }
                catch (IOException e)
                {
                    myErrors.Add("IOException for Field " + expectedHeader + " on row " + row + " : " + e.Message);
                }
            }
            return currRow.AddressLogId.Value;
        }
    }

}
