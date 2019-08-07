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
        static string DBString = ConfigurationManager.ConnectionStrings["DBString"].ConnectionString;
        static string DBTable = ConfigurationManager.AppSettings["dbTable"];

        static void Main(string[] args)
        {
            //DB connection
            SqlConnection conn = new SqlConnection(DBString);

            //Excel connection 
            DataTable Xdt = excelConnection();

            try
            {
                conn.Open();

                //clear existing data
                DeleteData(conn);
                PrintColourMessage(ConsoleColor.DarkGreen, "Table cleared");

                //Insert data
                SqlCommand insertCommand = SqlInsertStatement(Xdt, conn);
                connectToDB(insertCommand);
                PrintColourMessage(ConsoleColor.DarkGreen, "All rows inserted");

                //add a check to ensure each row is added
                //catch error, put all failed rows into a log.txt file
                //cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                PrintColourMessage(ConsoleColor.DarkRed, ex.Message + " Inner: " + ex.InnerException);
            }
            finally
            {
                conn.Close();
            }
            PrintColourMessage(ConsoleColor.White, "Press any key to exit");
            Console.ReadKey();
        }

        static void DeleteData(SqlConnection conn)
        {
            SqlCommand delTable = new SqlCommand($"DELETE FROM {DBTable}", conn);
            delTable.ExecuteNonQuery();
        }


        static DataTable excelConnection()
        {
            string ExcelFile = ConfigurationManager.ConnectionStrings["ExcelFile"].ConnectionString;
            OleDbConnection Xcon = new OleDbConnection(ExcelFile);
            //OleDbDataAdapter Xda = new OleDbDataAdapter("select * from [Sheet1$]", Xcon);
            DataTable Xdt = new DataTable();
            return Xdt;
        }

        static SqlCommand SqlInsertStatement(DataTable Xdt, SqlConnection conn)
        {
            int numberOfColumns = Xdt.Columns.Count;

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            string strInsertCommand = "";

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

                    //PrintColourMessage(ConsoleColor.Magenta, "Row: " + i + " " + lstRow[i].ToString() + " Type: " + lstRow[i].GetType());
                }
                string temp = lstRow[numberOfColumns - 1].ToString();
                lstRow[numberOfColumns - 1] = lstRow[numberOfColumns - 2];
                lstRow[numberOfColumns - 2] = temp;


                foreach (object item in lstRow)
                {
                    strInsertCommand += item + ", ";
                }
                strInsertCommand = strInsertCommand.Remove(strInsertCommand.Length - 2); //removes trailing space and comma
                PrintColourMessage(ConsoleColor.Yellow, strInsertCommand);
                Console.WriteLine(cmd.CommandText);

            }
            cmd.CommandText = $@"insert into {DBTable} VALUES({strInsertCommand})";
            return cmd;
        }


        static void connectToDB(SqlCommand cmd)
        {
            cmd.ExecuteNonQuery();
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


/*
    Connect to Excel spreadsheet / Get data out into DataTable - excelConnection()
    Manipulate data / prep insert statement - sqlStatement()
    Connect to Database / Delete current data in table / Upload new data - dbManipulate()
*/
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