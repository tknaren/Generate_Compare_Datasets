using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Xml;

using OfficeOpenXml;
using Generate_Compare_Datasets.Core;
using System.Xml.Serialization;

namespace Generate_Compare_Datasets.BusinessLogic
{
    internal class LoadData
    {
        const string MR_NAME_FIELD = "Name";
        const string MR_DESC_FIELD = "Description";
        //const string INPUT_FILE = @"C:\Naren\Working\Source\GitHubRepo\Generate_Compare_Datasets\MR_Data.xlsx";
        //const string OUTPUT_FILE = @"C:\Naren\Working\Source\GitHubRepo\Generate_Compare_Datasets\MR_Data.xml";

        /// <summary>
        /// Parent of Master method to start the program. Write the BL only inside other methods and call here
        /// </summary>
        public void GenerateXML(string inputFile, string outputFile)
        {
            DataSet dsExcelData = LoadDataSetFromExcel(inputFile);

            List<MRColumnDef> mdDefinitions = LoadMRTemplateDefinition(dsExcelData);

            DataSet dsData = RemoveDefinitionRowsFromTable(dsExcelData);

            MRTemplate mrObject = CreateMRObjectUsingDefinition(mdDefinitions, dsData);

            ConvertMRObjectToXml(mrObject, outputFile);

        }

        /// <summary>
        /// Load the data from Excel and populate the DataSet for easy manipulation
        /// </summary>
        /// <returns></returns>
        private DataSet LoadDataSetFromExcel(string inputFilePath)
        {
            //conf
            //string filePath = INPUT_FILE; // Replace with the path to your Excel file

            //conf
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Access the first worksheet

                DataTable dataTable = new DataTable();

                // Create table columns from the first row in the Excel file
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text);
                }

                // Add data rows to the DataTable
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var excelRow = worksheet.Cells[row, 1, row, worksheet.Dimension.End.Column];
                    var dataRow = dataTable.NewRow();
                    var endColumn = excelRow.End.Column;

                    for (int col = 1; col <= endColumn; col++)
                    {
                        dataRow[col - 1] = excelRow[row, col].Text;
                    }

                    dataTable.Rows.Add(dataRow);
                }

                // Create a DataSet and add the DataTable to it
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);

                // You have the Excel data in a DataSet
                return dataSet;
            }

        }
    
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dsExcelData"></param>
        private List<MRColumnDef> LoadMRTemplateDefinition(DataSet dsExcelData)
        {
            List<MRColumnDef> mRColumnDefs = new List<MRColumnDef>();

            DataTable dtSource = dsExcelData.Tables[0];

            for (int cnt=1; cnt < dtSource.Columns.Count; cnt++ )
            {
                MRColumnDef colDef = new MRColumnDef();

                colDef.nameInXML = dtSource.Columns[cnt].ColumnName;
                colDef.dataType = dtSource.Rows[0].Field<string>(cnt);
                colDef.pair = dtSource.Rows[1].Field<string>(cnt);
                colDef.nameInExcel = dtSource.Rows[2].Field<string>(cnt);
                colDef.order = cnt;

                mRColumnDefs.Add(colDef);
            }

            return mRColumnDefs;

        }

        /// <summary>
        /// 
        /// </summary>
        private DataSet RemoveDefinitionRowsFromTable(DataSet dsExcelData)
        {
            /**
             * Take a copy of the datatable and remove the definition rows from it. 
             * Those are no longer required
             * **/

            DataSet dsData = dsExcelData.Copy();

            DataTable dtData = dsData.Tables[0];

            //remove the first columns of the Dataset
            //config, decision and index
            dtData.Columns.RemoveAt(0);

            //remove the first 3 rows of the DataSet
            //config, decision and index
            dtData.Rows[0].Delete();
            dtData.Rows[0].Delete();
            dtData.Rows[0].Delete();
            dtData.AcceptChanges();

            return dsData;
        }


        /// <summary>
        /// 
        /// </summary>
        private MRTemplate CreateMRObjectUsingDefinition(List<MRColumnDef> mdDefinitions, DataSet dsData)
        {
            /**
             * Filter the dataset using the MR #
             * If there are multiple rows returned, 
             * Create a MR object using the definition
             * **/
            int mrCnt = 0;

            MRTemplate mrObject = new MRTemplate();

            mrObject.masterRecipes = new List<MasterRecipe>();

            DataTable dtMRData = dsData.Tables[0];

            List<MRColumnDef> mdOrderedColumnDefs = mdDefinitions.OrderBy(item => item.order).ToList();

            List<string> uniqueMRNames = GetUniqueValues(dtMRData, MR_NAME_FIELD);

            foreach (var mrName in uniqueMRNames)
            {
                var firstRow = dtMRData.AsEnumerable()
                    .Where(row => row.Field<string>(MR_NAME_FIELD) == mrName)
                    .FirstOrDefault();

                if (firstRow == null)
                    break;

                mrCnt++;

                MasterRecipe mr = new MasterRecipe();
                mr.tag = mrName;
                mr.description = firstRow.IsNull(MR_DESC_FIELD) ? string.Empty : firstRow[MR_DESC_FIELD].ToString();
                mr.version = mrCnt;

                mr.mrParams = new List<MRParam>();

                foreach (var colDef in mdOrderedColumnDefs)
                {
                    MRParam mrParam = new MRParam();
                    mrParam.tag = colDef.nameInXML;
                    //mrParam.value = firstRow.IsNull(colDef.nameInXML) ? string.Empty : firstRow[colDef.nameInXML].ToString();
                    mrParam.value = GetValueForAppropriateField(dtMRData,colDef);
                    mrParam.version = colDef.order;

                    mr.mrParams.Add(mrParam);
                }

                mrObject.masterRecipes.Add(mr);
            }

            return mrObject;
        }


        /// <summary>
        /// Get the Unique MRNames from the Data
        /// </summary>
        /// <param name="dtData"></param>
        /// <returns></returns>
        private List<string> GetUniqueValues(DataTable dtData, string fieldName)
        {
            var uniqueValues = dtData.AsEnumerable()
                .Where(row => !String.IsNullOrEmpty(row.Field<string>(fieldName)))
                .Select(row => row.Field<string>(fieldName))
                .Distinct();

            return uniqueValues.ToList<string>();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dtData"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        private string GetFirstValue(DataTable dtData, string fieldName)
        {
            var firstRow = dtData.AsEnumerable()
                .Select(row => row.Field<string>(fieldName))
                .FirstOrDefault();

            return firstRow != null ? firstRow.ToString() : string.Empty;
        }

        /// <summary>
        /// Traverses thru the data and performs string manipulations to return the appropriate value for a field
        /// </summary>
        /// <param name="dtMRData"></param>
        /// <param name="fieldName"></param>
        private string GetValueForAppropriateField(DataTable dtMRData, MRColumnDef colDef)
        {
            /**
             * Different logic for different data types
             * If String[], then club all the unique records into one  string separated by comma(,), and enclose it within {}
             * If String, then take only the first element, enclose it within "" 
             * If Number, then take only the first element
             * If Bool, then take only the first element
             * If Pair, then create a dictionary and add items, with primary as key and secondary as value
             * **/
            string returnValue = string.Empty;

            if (string.IsNullOrEmpty(colDef.pair))
            {
                switch (colDef.dataType)
                {
                    case "string[]":
                        var uniqueValues = GetUniqueValues(dtMRData, colDef.nameInXML);
                        returnValue = string.Format("{0}{1}{2}","{",string.Concat("\"", string.Join("\",\"", uniqueValues), "\""),"}");
                        break;
                    case "string":
                        var stringValue = GetFirstValue(dtMRData, colDef.nameInXML);
                        returnValue = string.Format("{0}{1}{0}", "\"", stringValue);
                        break;
                    case "number":
                        var numberValue = GetFirstValue(dtMRData, colDef.nameInXML);
                        returnValue = string.Format("{0}", numberValue);
                        break;
                    case "boolean":
                        var boolValue = GetFirstValue(dtMRData, colDef.nameInXML);
                        returnValue = string.Format("{0}", boolValue);
                        break;
                }
            }
            else
            {

            }

            return returnValue;
        }


        /// <summary>
        /// Serialize an Object to XML
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="filePath"></param>
        private void ConvertMRObjectToXml<T>(T obj, string outputFilePath)
        {
            //string filePath = OUTPUT_FILE; //Replace with the xml file of your choice

            XmlSerializer serializer = new XmlSerializer(typeof(T));

            using (TextWriter writer = new StreamWriter(outputFilePath))
            {
                serializer.Serialize(writer, obj);
            }
        }
    }
}
