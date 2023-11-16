// See https://aka.ms/new-console-template for more information

using System;
using Generate_Compare_Datasets.BusinessLogic;

class Program
{
    static void Main(string[] args)
    {
        // Check if there are any command line arguments
        //if (args.Length == 0)
        //{
        //    Console.WriteLine("No command line arguments provided.");
        //}
        //else
        //{
            Console.WriteLine("Command line arguments:");

            // Display each command line argument
            //for (int i = 0; i < args.Length; i++)
            //{
            //    Console.WriteLine($"Argument {i + 1}: {args[i]}");
            //}

            LoadData ld = new LoadData();
            ld.GenerateXML(args[0], args[1]);

        //}
    }
}





/*************************************

1. Have a basic template of XML
2. Read the data from the Excel
    2a. Filter the data using the MR Name 
    2b. Identify the datatype mentioned for the specific column
    2c. DEfine the logic to fill the data appropriately in the xml
3. Create nodes and fill in appropriately

Load the excel into a dataset
the first row bneing the datatype, 
create an .net object with name, datatupe, order, pair match, null allowed?, and other special properties
    name in excel
    name in xml
    data type
    order
    pair
    null allowed
get the data from the xcel and load it to the dataset
find out the uniqye MR ID in the dataset
filter those ids one by one and pick the records
find the datatype for it in the custom object
create the XML element, and assign the values 
while assiging the values, implement the formatting logic based on datatype, pair and null
save the xml

accept the other xml generated out of 

*************************************/
