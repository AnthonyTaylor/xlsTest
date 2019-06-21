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
            //DB connection
            string DBString = ConfigurationManager.ConnectionStrings["DBString"].ConnectionString;
            string DBTable = ConfigurationManager.AppSettings["dbTable"];
            SqlConnection conn = new SqlConnection(DBString);
            conn.Open();

            //Excel connection 
            string ExcelFile = ConfigurationManager.ConnectionStrings["ExcelFile"].ConnectionString;
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
                    List<object> lstRow = new List<object>(); ;

                    // go through each column in the row
                    for (int i = 0; i < numberOfColumns; i++)
                    {
                        if (i == (numberOfColumns - 1) || i == (numberOfColumns - 2))
                        {
                            lstRow.Add(row[i].ToString());
                        }
                        else
                        {
                            lstRow.Add(row[i]);
                        }

                        if (lstRow[i].GetType() == typeof(System.String))
                        {
                            lstRow[i] = "\'" + lstRow[i] + "\'";
                        }
                        if (lstRow[i].GetType() == typeof(System.DateTime))
                        {
                            // string startdate = DateTime.Parse("25/12/2008").ToString("yyyy-MM-dd");
                            // arrRow[i] = startdate;
                            lstRow[i] = Convert.ToString("\'2008-01-10\'");
                        }

                        PrintColourMessage(ConsoleColor.Magenta, "Row: " + i + " " + lstRow[i].ToString() + " Type: " + lstRow[i].GetType());
                    }
                    string temp = lstRow[numberOfColumns - 1].ToString();
                    lstRow[numberOfColumns - 1] = lstRow[numberOfColumns - 2];
                    lstRow[numberOfColumns - 2] = temp;

                    string strInsertCommand = "";
                    foreach (object item in lstRow)
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
        /// <summary>
        /// Outputs a coloured messsage to the console
        /// </summary>
        /// <param name="colour">ConsoleColor - colour for the message</param>
        /// <param name="message">String - message to be sent to the console</param>
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
///Location for the file to be dropped
///place for application to watch folder
///</future>

///<now>
///Check on number of rows in file and compare to numnber of rows in db at the end 
///on sucess, send notification
/// on failure send notificaation
///If total failure, reload the previous one
///</now>