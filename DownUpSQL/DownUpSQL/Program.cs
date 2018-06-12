using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Drawing;

namespace DownUpSQL
{
    class Program
    {

        static void Main(string[] args)
        {
            //sentinel variable
            bool cont = true;

            //choice variable
            byte choice;

            do
            {
                //prompt user
                Console.WriteLine("Welcome to the SQL Management App!");
                Console.WriteLine("How would you like to proceed?\n");
                Console.WriteLine("0 -- Quit");
                Console.WriteLine("1 -- Download to File");
                Console.WriteLine("2 -- Upload from File");
                Console.WriteLine("3 -- Enter single Database entry");
                Console.WriteLine();
                Console.Write("Please make a selection >> ");
                choice = GetChoice(Console.ReadLine());
                //Console.Write(choice.ToString());

                switch (choice)
                {
                    //user enters 0 -- Exit
                    case 0:
                        Console.WriteLine("Exiting program...Have a nice day.");
                        cont = false;
                        break;

                    case 1:
                        DownloadData();
                        break;

                    case 2:
                        UploadDataFile();
                        break;

                    case 3:
                        UploadDataLine();
                        break;

                    default:
                        Console.WriteLine("Please enter only the accepted values.\n");
                        break;
                }
            }
            while (cont);
        }

        static byte GetChoice(string x)
        {
            byte choice = 0;
            bool tryAgain = true;
            do
            {
                try
                {

                    choice = Byte.Parse(x);
                    tryAgain = false;
                    //return choice;

                }
                catch (FormatException fe)
                {
                    Console.WriteLine("Oops!...{0}\n", fe.Message);

                    //reprompt user
                    Console.Write("Please enter a valid value >> ");

                    x = Console.ReadLine();
                }
            }
            while (tryAgain);
            return choice;
        }

        static StreamReader openReadFile(string x)
        {
            StreamReader sr = null;
            bool tryAgain = true;
            do
            {
                try
                {
                    sr = new StreamReader(x);
                    tryAgain = false;
                }
                catch (FileNotFoundException fnfe)
                {
                    Console.WriteLine("Oops!...{0}", fnfe.Message);

                    Console.Write("Please enter a valid filepath >> ");

                    x = Console.ReadLine();
                }
            }
            while (tryAgain);
            return sr;
        }

        static DateTime checkDateTime(string x)
        {
            DateTime dt = DateTime.Today;
            bool tryAgain = true;
            do
            {
                try
                {
                    dt = DateTime.Parse(x);
                    tryAgain = false;

                }
                catch(FormatException fe)
                {
                    Console.WriteLine("Oops!...{0}", fe.Message);

                    Console.Write("Please enter a valid value >> ");

                    x = Console.ReadLine();
                }
            }
            while (tryAgain);
            return dt;
        }

        static void DownloadData()
        {
            string connStr = @"Data Source=USEM-527643\GENOTYPE;Initial Catalog=Test1;Integrated Security=True;User ID=Application;Password=w4rl0Ck5!";

            bool tryAgain = true;
            byte filetype;
            string filepath;
            do
            {
                //prompt user for filetype
                Console.WriteLine("What type of file would you like to write to?\n");
                Console.WriteLine("0 -- Back");
                Console.WriteLine("1 -- Comma Delimited (.csv)");
                Console.WriteLine("2 -- Excel Workbook (.xlsx)");
                Console.WriteLine();
                Console.Write("Please make a selection >> ");

                filetype = GetChoice(Console.ReadLine());

                switch (filetype)
                {
                    case 0:
                        Console.WriteLine("Back to Main Menu\n");
                        tryAgain = false;
                        break;

                    //user enters 1 -- .csv file
                    case 1:
                        Console.Write("Please enter the file path to download to >> ");
                        filepath = Console.ReadLine();

                        using (SqlConnection db = new SqlConnection(connStr))
                        {
                            db.Open();

                            Console.Write("Downloading data...");

                            const string SQL = "SELECT FName, LName, DOB, Years_Employed, Active_Employee FROM Employees";
                            using (SqlCommand sqlCommand = new SqlCommand(SQL, db))
                            {
                                using (SqlDataReader reader = sqlCommand.ExecuteReader())
                                {
                                    using (StreamWriter sw = new StreamWriter(filepath))
                                    {
                                        object[] output = new object[reader.FieldCount];

                                        for (int i = 0; i < reader.FieldCount; i++)
                                        {
                                            output[i] = reader.GetName(i);
                                        }

                                        sw.WriteLine(string.Join(",", output));

                                        while (reader.Read())
                                        {
                                            reader.GetValues(output);
                                            sw.WriteLine(string.Join(",", output));
                                        }
                                    }
                                    //status update
                                    Console.WriteLine("Done");
                                }
                            }
                        }
                        Console.WriteLine("Back to Main Menu\n");
                        tryAgain = false;
                        break;

                    //user enters 2 -- .xlsx file
                    case 2:
                        Console.Write("Please enter the file path to download to >> ");
                        filepath = Console.ReadLine();

                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            excel.Workbook.Worksheets.Add("Data");

                            var excelWorksheet = excel.Workbook.Worksheets["Data"];

                            List<string[]> headerRow = new List<string[]>()
                                                {
                                                    new string[] {"First Name", "Last Name", "DOB", "Years Employed", "Active Employee"}
                                                };
                            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                            excelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);

                            using (SqlConnection db = new SqlConnection(connStr))
                            {
                                db.Open();

                                const string SQL = "SELECT FName, LName, DOB, Years_Employed, Active_Employee FROM Employees";

                                //status update
                                Console.Write("Downloading data...");

                                using (SqlCommand sqlCommand = new SqlCommand(SQL, db))
                                {
                                    using (SqlDataReader reader = sqlCommand.ExecuteReader())
                                    {
                                        List<object[]> dataHolder = new List<object[]>();

                                        while (reader.Read())
                                        {
                                            object[] output = new object[reader.FieldCount];
                                            reader.GetValues(output);
                                            dataHolder.Add(output);
                                        }

                                        List<object[]> cellData = new List<object[]>();

                                        foreach (object[] x in dataHolder)
                                        {
                                            DateTime tempDOB = DateTime.Parse(x[2].ToString());
                                            x[2] = tempDOB.ToString("yyyy/MM/dd");

                                            cellData.Add(x);
                                        }

                                        excelWorksheet.Cells[2, 1].LoadFromArrays(cellData);

                                        var start = excelWorksheet.Dimension.Start;
                                        var end = excelWorksheet.Dimension.End;
                                        for (int x = start.Row; x <= end.Row; x++)
                                        {
                                            for(int i = start.Column; i <= end.Column; i++)
                                            {
                                                string cell = excelWorksheet.Cells[x, i].Address;
                                                string statement = "IF(ISNUMBER(" + cell + "),IF(AND(" + cell + ">=10," + cell + "<20),TRUE,FALSE),FALSE)";
                                                var cond1 = excelWorksheet.ConditionalFormatting.AddExpression(new ExcelAddress(cell));
                                                cond1.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                cond1.Style.Fill.BackgroundColor.Color = Color.Yellow;
                                                cond1.Formula = statement;

                                                statement = "IF(ISNUMBER(" + cell + "),IF(" + cell + "<10,TRUE,FALSE),FALSE)";
                                                var cond2 = excelWorksheet.ConditionalFormatting.AddExpression(new ExcelAddress(cell));
                                                cond2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                cond2.Style.Fill.BackgroundColor.Color = Color.Red;
                                                cond2.Formula = statement;

                                                statement = "IF(ISNUMBER(" + cell + "),IF(" + cell + ">=20,TRUE,FALSE),FALSE)";
                                                var cond3 = excelWorksheet.ConditionalFormatting.AddExpression(new ExcelAddress(cell));
                                                cond3.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                cond3.Style.Fill.BackgroundColor.Color = Color.Green;
                                                cond3.Formula = statement;

                                                statement = "EXACT(" + cell + ",TRUE)";
                                                var cond4 = excelWorksheet.ConditionalFormatting.AddExpression(new ExcelAddress(cell));
                                                cond4.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                cond4.Style.Fill.BackgroundColor.Color = Color.Green;
                                                cond4.Formula = statement;

                                                statement = "EXACT(" + cell + ",FALSE)";
                                                var cond5 = excelWorksheet.ConditionalFormatting.AddExpression(new ExcelAddress(cell));
                                                cond5.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                cond5.Style.Fill.BackgroundColor.Color = Color.Red;
                                                cond5.Formula = statement;
                                            }
                                        }
                                    }
                                }
                            }
                            //status update
                            Console.WriteLine("Done");

                            FileInfo excelFile = new FileInfo(@filepath);
                            excel.SaveAs(excelFile);
                        }
                        Console.WriteLine("Back to Main Menu\n");
                        tryAgain = false;
                        break;

                    default:
                        Console.WriteLine("Please enter only '1' or '2' for filetype selection\n");
                        break;
                }
            }
            while (tryAgain);
        }

        static void UploadDataFile()
        {
            string connStr = @"Data Source=USEM-527643\GENOTYPE;Initial Catalog=Test1;Integrated Security=True;User ID=Application;Password=w4rl0Ck5!";
            string fname;
            string lname;
            DateTime dob;
            int yearsEmployed;
            byte is_active;

            byte hasDuplicate = 2;
            bool tryAgain = true;
            byte filetype;
            string filepath;

            string temp;

            do
            {
                //prompt user for filetype
                Console.WriteLine("What type of file would you like to upload from?\n");
                Console.WriteLine("0 -- Back");
                Console.WriteLine("1 -- Comma Delimited (.csv)");
                Console.WriteLine("2 -- Excel Workbook (.xlsx)");
                Console.WriteLine();
                Console.Write("Please make a selection >> ");

                filetype = GetChoice(Console.ReadLine());

                switch (filetype)
                {
                    case 0:
                        tryAgain = false;
                        break;

                    //user enters 1 -- .csv file
                    case 1:
                        Console.Write("Does the file contain data that is already in the database? (y/n) >> ");
                        temp = Console.ReadLine();

                        do
                        {
                            if (temp == "y" || temp == "Y" || temp == "yes" || temp == "Yes" || temp == "YES")
                            {
                                hasDuplicate = 1;
                            }
                            else if (temp == "n" || temp == "N" || temp == "no" || temp == "No" || temp == "NO")
                            {
                                hasDuplicate = 0;
                            }
                            else
                            {
                                Console.Write("Please enter a valid value >> ");
                                temp = Console.ReadLine();
                            }
                        }
                        while (hasDuplicate == 2);

                        switch (hasDuplicate)
                        {
                            case 0:
                                using (SqlConnection db = new SqlConnection(connStr))
                                {
                                    Console.Write("Enter the file path of the .csv file containing the data to upload >> ");
                                    filepath = Console.ReadLine();
                                    StreamReader sr = openReadFile(filepath);

                                    db.Open();

                                    //status update
                                    Console.Write("Reading Data...");

                                    string line = sr.ReadLine();
                                    string[] value = line.Split(',');

                                    //status update
                                    Console.Write("Uploading Data...");

                                    while (!sr.EndOfStream)
                                    {
                                        value = sr.ReadLine().Split(',');

                                        if (value.Length == 5)
                                        {
                                            //convert data types to appropriate ones
                                            fname = value[0];
                                            lname = value[1];
                                            dob = DateTime.Parse(value[2]);
                                            yearsEmployed = Int32.Parse(value[3]);

                                            if (value[4] == "TRUE" || value[4] == "True" || value[4] == "true")
                                            {
                                                is_active = 1;
                                            }
                                            else
                                            {
                                                is_active = 0;
                                            }

                                            //build command string
                                            string s = "INSERT INTO Employees (FName, LName, DOB, Years_Employed, Active_Employee) VALUES (@fname, @lname, @dob, @years, @active)";
                                            using (SqlCommand sql = new SqlCommand(s, db))
                                            {
                                                sql.Parameters.Add("@fname", SqlDbType.VarChar).Value = fname;
                                                sql.Parameters.Add("@lname", SqlDbType.VarChar).Value = lname;
                                                sql.Parameters.Add("@dob", SqlDbType.Date).Value = dob;
                                                sql.Parameters.Add("@years", SqlDbType.Int).Value = yearsEmployed;
                                                sql.Parameters.Add("@active", SqlDbType.Bit).Value = is_active;

                                                sql.ExecuteNonQuery();
                                            }
                                        }
                                    }
                                    sr.Close();
                                    //status update
                                    Console.WriteLine("Done");
                                }
                                break;

                            case 1:
                                using (SqlConnection db = new SqlConnection(connStr))
                                {
                                    Console.Write("Enter the file path of the .csv file containing the data to upload >> ");
                                    filepath = Console.ReadLine();
                                    StreamReader sr = openReadFile(filepath);

                                    db.Open();
                                    SqlCommand trunc = new SqlCommand("TRUNCATE TABLE Employees", db);
                                    trunc.ExecuteNonQuery();
                                    trunc.Dispose();

                                    //status update
                                    Console.Write("Reading Data...");

                                    string line = sr.ReadLine();
                                    string[] value = line.Split(',');

                                    //status update
                                    Console.Write("Uploading Data...");

                                    while (!sr.EndOfStream)
                                    {
                                        value = sr.ReadLine().Split(',');

                                        if (value.Length == 5)
                                        {
                                            //convert data types to appropriate ones
                                            fname = value[0];
                                            lname = value[1];
                                            dob = DateTime.Parse(value[2]);
                                            yearsEmployed = Int32.Parse(value[3]);

                                            if (value[4] == "TRUE" || value[4] == "True" || value[4] == "true")
                                            {
                                                is_active = 1;
                                            }
                                            else
                                            {
                                                is_active = 0;
                                            }

                                            //build command string
                                            string s = "INSERT INTO Employees (FName, LName, DOB, Years_Employed, Active_Employee) VALUES (@fname, @lname, @dob, @years, @active)";
                                            using (SqlCommand sql = new SqlCommand(s, db))
                                            {
                                                sql.Parameters.Add("@fname", SqlDbType.VarChar).Value = fname;
                                                sql.Parameters.Add("@lname", SqlDbType.VarChar).Value = lname;
                                                sql.Parameters.Add("@dob", SqlDbType.Date).Value = dob;
                                                sql.Parameters.Add("@years", SqlDbType.Int).Value = yearsEmployed;
                                                sql.Parameters.Add("@active", SqlDbType.Bit).Value = is_active;

                                                sql.ExecuteNonQuery();
                                            }
                                        }
                                    }
                                    sr.Close();
                                    //status update
                                    Console.WriteLine("Done");
                                }
                                break;
                        }

                        tryAgain = false;
                        break;

                    //user enters 2 -- .xlsx file
                    case 2:
                        Console.Write("Does the file contain data that is already in the database? (y/n) >> ");
                        temp = Console.ReadLine();

                        do
                        {
                            if (temp == "y" || temp == "Y" || temp == "yes" || temp == "Yes" || temp == "YES")
                            {
                                hasDuplicate = 1;
                            }
                            else if (temp == "n" || temp == "N" || temp == "no" || temp == "No" || temp == "NO")
                            {
                                hasDuplicate = 0;
                            }
                            else
                            {
                                Console.Write("Please enter a valid value >> ");
                                temp = Console.ReadLine();
                            }
                        }
                        while (hasDuplicate == 2);

                        switch (hasDuplicate)
                        {
                            case 0:
                                using (SqlConnection db = new SqlConnection(connStr))
                                {
                                    Console.Write("Enter the file path of the .csv file containing the data to upload >> ");
                                    filepath = Console.ReadLine();
                                    using (ExcelPackage excel = new ExcelPackage(new FileInfo(filepath)))
                                    {
                                        //status update
                                        Console.Write("Reading Data...");

                                        var myWorksheet = excel.Workbook.Worksheets.First();
                                        var totalRows = myWorksheet.Dimension.End.Row;
                                        var totalColumns = myWorksheet.Dimension.End.Column;

                                        //status update
                                        Console.Write("Uploading Data...");

                                        db.Open();

                                        StringBuilder sb = new StringBuilder();
                                        for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                                        {
                                            var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                                            sb.AppendLine(string.Join(",", row));

                                            string line = sb.ToString();

                                            string[] value = line.Split(',');

                                            //convert data types to appropriate ones
                                            fname = value[0];
                                            lname = value[1];
                                            dob = DateTime.Parse(value[2]);
                                            yearsEmployed = Int32.Parse(value[3]);

                                            if (value[4] == "TRUE" || value[4] == "True" || value[4] == "true")
                                            {
                                                is_active = 1;
                                            }
                                            else
                                            {
                                                is_active = 0;
                                            }


                                            //build command string
                                            string s = "INSERT INTO Employees (FName, LName, DOB, Years_Employed, Active_Employee) VALUES (@fname, @lname, @dob, @years, @active)";
                                            using (SqlCommand sql = new SqlCommand(s, db))
                                            {
                                                sql.Parameters.Add("@fname", SqlDbType.VarChar).Value = fname;
                                                sql.Parameters.Add("@lname", SqlDbType.VarChar).Value = lname;
                                                sql.Parameters.Add("@dob", SqlDbType.Date).Value = dob;
                                                sql.Parameters.Add("@years", SqlDbType.Int).Value = yearsEmployed;
                                                sql.Parameters.Add("@active", SqlDbType.Bit).Value = is_active;

                                                sql.ExecuteNonQuery();
                                            }
                                        }
                                    }
                                    //status update
                                    Console.WriteLine("Done");
                                }
                                break;

                            case 1:
                                using (SqlConnection db = new SqlConnection(connStr))
                                {
                                    Console.Write("Enter the file path of the .csv file containing the data to upload >> ");
                                    filepath = Console.ReadLine();
                                    using (ExcelPackage excel = new ExcelPackage(new FileInfo(filepath)))
                                    {
                                        //status update
                                        Console.Write("Reading Data...");

                                        var myWorksheet = excel.Workbook.Worksheets.First();
                                        var totalRows = myWorksheet.Dimension.End.Row;
                                        var totalColumns = myWorksheet.Dimension.End.Column;

                                        //status update
                                        Console.Write("Uploading Data...");

                                        db.Open();
                                        SqlCommand trunc = new SqlCommand("TRUNCATE TABLE Employees", db);
                                        trunc.ExecuteNonQuery();
                                        trunc.Dispose();

                                        StringBuilder sb = new StringBuilder();
                                        for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                                        {
                                            var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                                            sb.AppendLine(string.Join(",", row));

                                            string line = sb.ToString();

                                            string[] value = line.Split(',');

                                            //convert data types to appropriate ones
                                            fname = value[0];
                                            lname = value[1];
                                            dob = DateTime.Parse(value[2]);
                                            yearsEmployed = Int32.Parse(value[3]);

                                            if (value[4] == "TRUE" || value[4] == "True" || value[4] == "true")
                                            {
                                                is_active = 1;
                                            }
                                            else
                                            {
                                                is_active = 0;
                                            }


                                            //build command string
                                            string s = "INSERT INTO Employees (FName, LName, DOB, Years_Employed, Active_Employee) VALUES (@fname, @lname, @dob, @years, @active)";
                                            using (SqlCommand sql = new SqlCommand(s, db))
                                            {
                                                sql.Parameters.Add("@fname", SqlDbType.VarChar).Value = fname;
                                                sql.Parameters.Add("@lname", SqlDbType.VarChar).Value = lname;
                                                sql.Parameters.Add("@dob", SqlDbType.Date).Value = dob;
                                                sql.Parameters.Add("@years", SqlDbType.Int).Value = yearsEmployed;
                                                sql.Parameters.Add("@active", SqlDbType.Bit).Value = is_active;

                                                sql.ExecuteNonQuery();
                                            }
                                        }
                                    }
                                    //status update
                                    Console.WriteLine("Done");
                                }
                                break;
                        }

                        tryAgain = false;
                        break;
                }
            }
            while (tryAgain);
        }

        static void UploadDataLine()
        {
            string connStr = @"Data Source=USEM-527643\GENOTYPE;Initial Catalog=Test1;Integrated Security=True;User ID=Application;Password=w4rl0Ck5!";

            string fname;
            string lname;
            DateTime dob;
            byte yearsEmployed;
            byte is_active = 3;

            Console.Write("Please enter the First Name of the employee >> ");

            fname = Console.ReadLine();
            if (fname.Contains(" ") || fname.Contains("\t") || fname == "")
            {
                do
                {
                    Console.Write("Please enter a valid value >> ");
                    fname = Console.ReadLine();
                }
                while (fname.Contains(" ") || fname.Contains("\t") || fname == "");
            }

            Console.Write("Please enter the Last Name of the employee >> ");
            lname = Console.ReadLine();
            if (lname.Contains(" ") || lname.Contains("\t") || lname == "")
            {
                do
                {
                    Console.Write("Please enter a valid value >> ");
                    lname = Console.ReadLine();
                }
                while (lname.Contains(" ") || lname.Contains("\t") || lname == "");
            }

            Console.Write("Please enter the Date of Birth of the employee (yyyy/MM/dd) >> ");
            dob = checkDateTime(Console.ReadLine());

            Console.Write("Please enter the number of years the employee has been employed >> ");
            yearsEmployed = GetChoice(Console.ReadLine());

            Console.Write("Please enter whether the employee is acitve (y/n) >> ");
            string tempAct = Console.ReadLine();
            do
            {
                if (tempAct == "y" || tempAct == "Y" || tempAct == "yes" || tempAct == "Yes" || tempAct == "YES")
                {
                    is_active = 1;
                }
                else if (tempAct == "n" || tempAct == "N" || tempAct == "no" || tempAct == "No" || tempAct == "NO")
                {
                    is_active = 0;
                }
                else
                {
                    Console.WriteLine("Please enter a valid value >> ");
                    tempAct = Console.ReadLine();
                }
            }
            while (is_active == 3);

            Console.WriteLine("Please verify that the information is correct\n");
            Console.WriteLine("First Name -- " + fname);
            Console.WriteLine("Last Name -- " + lname);
            Console.WriteLine("Date of Birth -- " + dob);
            Console.WriteLine("Years Employed -- " + yearsEmployed);
            if (is_active == 1)
            {
                Console.WriteLine("Is Active -- YES\n");
            }
            else
            {
                Console.WriteLine("Is Active -- NO\n");
            }

            Console.Write("Is the information correct? (y/n) >> ");
            tempAct = Console.ReadLine();
            if (tempAct == "y" || tempAct == "Y" || tempAct == "yes" || tempAct == "Yes" || tempAct == "YES")
            {
                Console.Write("Uploading...");

                using (SqlConnection db = new SqlConnection(connStr))
                {
                    db.Open();
                    string s = "INSERT INTO Employees (FName, LName, DOB, Years_Employed, Active_Employee) VALUES (@fname, @lname, @dob, @years, @active)";
                    using (SqlCommand sql = new SqlCommand(s, db))
                    {
                        sql.Parameters.Add("@fname", SqlDbType.VarChar).Value = fname;
                        sql.Parameters.Add("@lname", SqlDbType.VarChar).Value = lname;
                        sql.Parameters.Add("@dob", SqlDbType.Date).Value = dob;
                        sql.Parameters.Add("@years", SqlDbType.Int).Value = yearsEmployed;
                        sql.Parameters.Add("@active", SqlDbType.Bit).Value = is_active;

                        sql.ExecuteNonQuery();

                        Console.WriteLine("Done");
                    }
                }
            }
            Console.WriteLine("Back to Main Menu\n");
        }
    }
}