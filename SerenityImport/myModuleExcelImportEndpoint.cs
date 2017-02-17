namespace myProject.Default.Entities
{
    using Repositories;
    using OfficeOpenXml;
    using Serenity;
    using Serenity.Data;
    using Serenity.Services;
    using Serenity.Web;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Web.Mvc;
    using MyRow = CustomerRow;
    using System.Linq;
    using myImpHelp = myImportHelper.ExcelImportHelper;
    using myImpHelpExt = myImportHelper.myExtension;
    using jImpHelp = myImportJointFields.ExcelImportHelper;


    [RoutePrefix("Services/Default/CustomerExcelImport"), Route("{action}")]
    [ConnectionKey(typeof(MyRow)), ServiceAuthorize]
    public class CustomerExcelImportController : ServiceEndpoint
    {
        [HttpPost]
        public ExcelImportResponse ExcelImport(IUnitOfWork uow, ExcelImportRequest request)
        {
            //-------------------------- Gather Excel File Data ------------------------------------------------------------//

            request.CheckNotNull();
            var fName = request.FileName;
            Check.NotNullOrWhiteSpace(fName, "filename");
            UploadHelper.CheckFileNameSecurity(fName);

            if (!request.FileName.StartsWith("temporary/"))
                throw new ArgumentOutOfRangeException("filename");

            ExcelPackage ep = new ExcelPackage();
            using (var fs = new FileStream(UploadHelper.DbFilePath(fName), FileMode.Open, FileAccess.Read))
                ep.Load(fs);

            var response = new ExcelImportResponse();
            var myErrors = response.ErrorList = new List<string>();

            /*Read the first Excell sheet and then gather the headers of the import file*/
            var worksheet = ep.Workbook.Worksheets.First();
            var wsStart = worksheet.Dimension.Start;
            var wsEnd = worksheet.Dimension.End;
            var headers = worksheet.Cells[wsStart.Row, wsStart.Column, 1, wsEnd.Column];

            //-------------------------- Gather Mapping nformation ------------------------------------------------------------//
                        
            /*A few variables to make our life easier*/
            var myConnection = uow.Connection;
            var myFields = MyRow.Fields; 
            List<string> importedHeaders = new List<string>(); //Headers from Imported File 
            List<object> importedValues = new List<object>(); //Values being Imported
            List<string> systemHeaders = new List<string>(); //Headers currently in system
            List<string> sysHeader = new List<string>(); //System Header to import
            List<string> exceptionHeaders = new List<string>(); //Haders to not check for during import. 
            object obj = null; //Object container for value being imported
            dynamic a = null; //Handled object to assign to system
            string fieldTitle = ""; //Title of field being imported
            jImpHelp.entryType entType; //Type of handler to use. 
            
            /*Add Imported file headers to proper list*/
            foreach (var q in headers)
            {
                importedHeaders.Add(q.Text);
            }            

            /*  Add system headers to proper list while also adding 'ID' to the list. 'ID'
             *  is the key field from exported files and needs to be mapped manually */            
            systemHeaders.Add("ID");
            foreach (var t in myFields)
            {
                systemHeaders.Add(t.Title);                
            };

            /* Not all columns will be expected to be imported. To avoid unnecesary error messages 
             * we add the titles of the fields we want ignored here.*/
            exceptionHeaders.Add(myFields.AddressLogId.Title);


            /* Using the systemHeaders to compare against the importedHeaders, we build an index with 
             * the column location and match it to the system header using a Dictionary<string, int>. */

            Dictionary<string, int> headerMap = myImpHelp.myExcelHeaderMap(importedHeaders, systemHeaders, myErrors, exceptionHeaders);           

            for (var row = 2; row <= wsEnd.Row; row++)
            {               
                try
                {
                    /* This instance checks the ID field as to whether the row exists or not. if the 
                     * ID key exists, it will use it to update the row with the imported fields but if
                     * it does not exist, it creates a new entry. */
                                         
                    var sysKey = myFields.CustomerId; 
                    obj = myImpHelp.myExcelVal(row, myImpHelpExt.GetEntry(headerMap, "ID").Value, worksheet);
                    var wsKeyField = Convert.ToInt32(obj);
                    var currentRow = myConnection.TryFirst<MyRow>(q => q.Select(sysKey).Where(sysKey == wsKeyField));

                    if (currentRow == null) //Create New if Row doesnt' exist
                        currentRow = new MyRow() { }; 
                    else
                        currentRow.TrackWithChecks = false;

                    /* We now need to handle how we want to manage the imported fields. We list the fields 
                     * being imported using the same code set but we update the entType, fieldTitle and then 
                     * designate what field will be updated. You handle specialty case handlers such as joint 
                     * fields in another file. Do note that I use importedValues and sysHeader to pass on the 
                     * values to the handler. I did this on purpose so that you can pass multiple values if 
                     * neccesary. For example, if your joint field requires 2 values to be entered, you can 
                     * simply capture the value with an additional .Add()
                       
                    Example : 
                    
                    -------Simple Field
                    entType = jImpHelp.entryType.String; //designate the type of item
                    fieldTitle = myFields.firstField.Title; //designate the field to be looked at
                    obj = myImpHelp.myExcelVal(row, myImpHelpExt.GetEntry(headerMap, fieldTitle).Value, worksheet);
                    ----ADD OBJECT TO NEW CONTAINER--
                    var newContainer = obj;
                    ----------------------
                   if (obj != null)
                    {
                        importedValues.Add(obj);
                        sysHeader.Add(fieldTitle);
                        a = jImpHelp.myImportEntry(importedValues, myErrors, sysHeader, row, entType, myConnection);
                        if (a != null)
                        {
                            currentRow.CustomerAddress = a; //designate the field to be updated in the system
                        }
                        sysHeader.Clear();
                        importedValues.Clear();
                    }
                    -----Joint Field
                    entType = jImpHelp.entryType.jointField; //<--Update Me according to type of field to merge with
                    fieldTitle = myFields.secondFieldCustomerName.Title;//<--Update Me
                    obj = myImpHelp.myExcelVal(row, myImpHelpExt.GetEntry(headerMap, fieldTitle).Value, worksheet);
                    if (obj != null)
                    {                    
                        importedValues.Add(obj);
                        -----ADD CAPTURED CONTAINER TO VALUES LIST-----                        
                        importedValue.Add(newContainer)
                        -------------------------------------------                                      
                        sysHeader.Add(fieldTitle);
                        a = jImpHelp.myImportEntry(importedValues, myErrors, sysHeader, row, entType, myConnection);
                        if (a != null)
                        {
                            currentRow.AddressLogId = a; ////<--Update Me. *Special Case. Notice: Not the same as field to match
                        }
                        sysHeader.Clear();
                        importedValues.Clear();
                    }
                    */

                    //--------------------------Merge Imported Fields ------------------------------------------------------------//
                                       
                    entType = jImpHelp.entryType.String; //designate the type of item
                    fieldTitle = myFields.CustomerAddress.Title; //designate the field to be looked at
                    obj = myImpHelp.myExcelVal(row, myImpHelpExt.GetEntry(headerMap, fieldTitle).Value, worksheet);
                    if (obj != null)
                    {
                        importedValues.Add(obj);
                        sysHeader.Add(fieldTitle);
                        a = jImpHelp.myImportEntry(importedValues, myErrors, sysHeader, row, entType, myConnection);
                        if (a != null)
                        {
                            currentRow.CustomerAddress = a; //designate the field to be updated in the system
                        }
                        sysHeader.Clear();
                        importedValues.Clear();
                    }

                    /*Same as above, just updated for the next field. */
                    entType = jImpHelp.entryType.String; //<--Update Me according to type of field to merge with
                    fieldTitle = myFields.CustomerName.Title;//<--Update Me
                    obj = myImpHelp.myExcelVal(row, myImpHelpExt.GetEntry(headerMap, fieldTitle).Value, worksheet);
                    if (obj != null)
                    {
                        importedValues.Add(obj);
                        sysHeader.Add(fieldTitle);
                        a = jImpHelp.myImportEntry(importedValues, myErrors, sysHeader, row, entType, myConnection);
                        if (a != null)
                        {
                            currentRow.CustomerName = a; //<--Update Me
                        }
                        sysHeader.Clear();
                        importedValues.Clear();
                    }

                    entType = jImpHelp.entryType.CategoryJoin; //<--Update Me *Special Case Joint Field
                    fieldTitle = myFields.AddressFloor.Title; //<--Update Me *Special Case. Notice: Field to match in import file
                    obj = myImpHelp.myExcelVal(row, myImpHelpExt.GetEntry(headerMap, fieldTitle).Value, worksheet);                    
                    if (obj != null)
                    {
                        importedValues.Add(obj);
                        sysHeader.Add(fieldTitle);
                        a = jImpHelp.myImportEntry(importedValues, myErrors, sysHeader, row, entType, myConnection);
                        if (a != null)
                        {
                            currentRow.AddressLogId = a; ////<--Update Me. *Special Case. Notice: Not the same as field to match
                        }
                        sysHeader.Clear();
                        importedValues.Clear();
                    }
                    //----------------------------------------Run Object Entries with Create or Update ------------------------------------//
                    if (currentRow.CustomerId == null)
                    {
                        new CustomerRepository().Create(uow, new SaveWithLocalizationRequest<MyRow>
                        {
                            Entity = currentRow
                        });
                        response.Inserted = response.Inserted + 1;
                    }
                    else
                    {
                        new CustomerRepository().Update(uow, new SaveWithLocalizationRequest<MyRow>
                        {
                            Entity = currentRow,
                            EntityId = currentRow.CustomerId.Value
                        });
                        response.Updated = response.Updated + 1;
                    }
                }
                catch (Exception ex)
                {
                    myErrors.Add(myImpHelp.eMessage3(row, ex.Message));
                }
            }
            return response;
        }        
    }
}
