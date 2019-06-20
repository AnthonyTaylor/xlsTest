using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections.Generic;
using System.Globalization;

namespace xlsxToDataBase
{
    class Program
    {
        static void Main(string[] args)
        {
            string DBString = ConfigurationManager.ConnectionStrings["DBString"].ConnectionString;
            string DBTable = ConfigurationManager.AppSettings["dbTable"];
            string ExcelFile = ConfigurationManager.ConnectionStrings["ExcelFile"].ConnectionString;

            //SqlConnection conn = new SqlConnection(@"data source=localhost;Server=.\SQLEXPRESS;Database=DBConnect;Integrated Security=SSPI;");
            SqlConnection conn = new SqlConnection(DBString);
            conn.Open();
            OleDbConnection Xcon = new OleDbConnection(ExcelFile);
            OleDbDataAdapter Xda = new OleDbDataAdapter("select * from [Sheet1$]", Xcon);
            DataTable Xdt = new DataTable();
            Xda.Fill(Xdt);
            try
            {
                SqlCommand delTable = new SqlCommand($"DELETE FROM {DBTable}", conn);
                delTable.ExecuteNonQuery();

                int numberOfColumns = Xdt.Columns.Count;

                foreach (DataRow row in Xdt.Rows) // Loop over the rows.
                {
                    //loop through each column here to add the values to an array
                    //swap the last two items in the array
                    //send array 
                    List<object> arrRow = new List<object>(); ;

                    // go through each column in the row
                    for (int i = 0; i < numberOfColumns; i++)
                    {
                        if (i == (numberOfColumns - 1) || i == (numberOfColumns - 2))
                        {
                            arrRow.Add(row[i].ToString());
                        }
                        else
                        {
                            arrRow.Add(row[i]);
                        }

                        if (arrRow[i].GetType() == typeof(System.String))
                        {
                            arrRow[i] = "\'" + arrRow[i] + "\'";
                        }
                        if (arrRow[i].GetType() == typeof(System.DateTime))
                        {
                            // string startdate = DateTime.Parse("25/12/2008").ToString("yyyy-MM-dd");
                            // arrRow[i] = startdate;
                            arrRow[i] = Convert.ToString("\'2008-01-10\'");
                        }

                        PrintColourMessage(ConsoleColor.Magenta, "Row: " + i + " " + arrRow[i].ToString() + " Type: " + arrRow[i].GetType());
                    }
                    string temp = arrRow[numberOfColumns - 1].ToString();
                    arrRow[numberOfColumns - 1] = arrRow[numberOfColumns - 2];
                    arrRow[numberOfColumns - 2] = temp;

                    string strInsertCommand = "";
                    foreach (object item in arrRow)
                    {
                        strInsertCommand += item + ", ";
                    }
                    strInsertCommand = strInsertCommand.Remove(strInsertCommand.Length - 2); //removes trailing space and comma
                    PrintColourMessage(ConsoleColor.Yellow, strInsertCommand);
                    SqlCommand cmd = new SqlCommand($@"insert into {DBTable} VALUES({strInsertCommand})", conn);
                    Console.WriteLine(cmd.CommandText);

                    //add a check to ensure each row is added
                    //catch error, put all failed rows into a log.txt file
                    cmd.ExecuteNonQuery();
                }
                PrintColourMessage(ConsoleColor.DarkGreen, "Number of rows added: " + Xdt.Rows.Count);
            }
            catch (Exception ex) { PrintColourMessage(ConsoleColor.DarkRed, ex.Message + " Inner: " + ex.InnerException); }
            finally { conn.Close(); }
            PrintColourMessage(ConsoleColor.White, "Press any key to exit");
            Console.ReadKey();
        }
        //Print colour message and reset colour
        static void PrintColourMessage(ConsoleColor colour, string message)
        {
            //Change text colour
            Console.ForegroundColor = colour;
            Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}

///<future>
//Location for the file to be dropped
//place for application to watch folder
///</future>

//<now>
//Check on number of rows in file and compare to numnber of rows in db at the end 
//on sucess, send notification
// on failure send notificaation
//If total failure, reload the previous one
///</now>