using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;
using ClosedXML.Excel;

namespace DBResultExcelSheet
{
    class DBHelper
    {
        private static readonly string connectionString = ConfigurationManager.ConnectionStrings["StudentDB"].ConnectionString;

        public static int dbHelperRowCount = 1;

        public void StoreProcedureQuery(string storeProcedureName, params SqlParameter[] parameters)
        {

            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                sqlConnection.Open();
                SqlCommand sqlCommand = new SqlCommand(storeProcedureName, sqlConnection);
                sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;

                /*if (parameters != null && parameters.Length > 0)
                    sqlCommand.Parameters.AddRange(parameters);*/

                CreateExcelFile(parameters);
                //Console.WriteLine(parameters[0].Value);

                /*sqlCommand.ExecuteNonQuery();
                //Console.WriteLine(numberOfRowAffected);
                sqlConnection.Close();*/
            }
        }


        private static void CreateExcelFile(params SqlParameter[] parameters)
        {
            //dbHelperRowCount++;

            using (var workbook = new XLWorkbook("CourseCodeResultSheet.xlsx"))
            {
                

                /*if (workbook.Worksheets.Contains("CourseCodeResultSheet"))
                {
                    var worksheet = workbook.Worksheets.Worksheet("CourseCodeResultSheet");
                    Console.WriteLine("if" + workbook.Worksheets.Contains("CourseCodeResultSheet"));
                }
                else
                {
                    Console.WriteLine("Else" + workbook.Worksheets.Contains("CourseCodeResultSheet"));
                    var worksheet = workbook.Worksheets.Add("CourseCodeResultSheet");

                }*/

                var worksheet = workbook.Worksheets.Worksheet("CourseCodeResultSheet");

                if (dbHelperRowCount == 1)
                {
                    worksheet.Cell(1, 1).Value = "Course Code";
                    worksheet.Cell(1, 2).Value = "Course Name";
                    worksheet.Cell(1, 3).Value = "Level";
                    worksheet.Cell(1, 4).Value = "Start Date";
                    worksheet.Cell(1, 5).Value = "Is Active";
                    worksheet.Cell(1, 6).Value = "Result";
                    worksheet.Cell(1, 7).Value = "Description";
                    dbHelperRowCount++;
                }

                for (int i = 0, j = 1; i < parameters.Length; i++, j++)
                {
                    switch (i)
                    {
                        case 0:
                            if (string.IsNullOrEmpty((string)parameters[i].Value))
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = "";
                                worksheet.Cell(dbHelperRowCount, 7).Value = "Course Code can't be empty";
                                break;
                            }
                            else
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = parameters[i].Value;
                            }
                            break;
                        case 1:
                            //string str = (string)parameters[i].Value;
                            if (string.IsNullOrEmpty((string)parameters[i].Value))
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = "";
                                worksheet.Cell(dbHelperRowCount, 7).Value = worksheet.Cell(dbHelperRowCount, 7).Value + ". " + "Course Name can't be empty";
                                break;
                            }
                            else
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = parameters[i].Value;
                            }
                            break;

                        case 2:
                            int checkInteger = 0;
                            string s = parameters[i].Value + "";
                            bool result = int.TryParse(s, out checkInteger);
                            if (result)
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = parameters[i].Value;
                            }
                            else
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = "";
                                worksheet.Cell(dbHelperRowCount, 7).Value = worksheet.Cell(dbHelperRowCount, 7).Value + ". " + "Level must be integer";
                            }
                            break;
                        case 3:
                            var dateTimeCheck = parameters[i].Value as DateTime?;

                            if (dateTimeCheck == null)
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = "";
                                worksheet.Cell(dbHelperRowCount, 7).Value = worksheet.Cell(dbHelperRowCount, 7).Value + ". " + "Date formet is not right";
                            }
                            else
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = parameters[i].Value;
                            }
                            break;
                        case 4:
                            var isActive = parameters[i].Value as bool?;
                            if (isActive.HasValue)
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = parameters[i].Value;
                            }
                            else
                            {
                                worksheet.Cell(dbHelperRowCount, j).Value = "";
                                worksheet.Cell(dbHelperRowCount, 7).Value = worksheet.Cell(dbHelperRowCount, 7).Value + ". " + "Course must be active";
                            }
                            break;
                    }
                    Console.WriteLine("Row Number : "  + dbHelperRowCount + ": " +worksheet.Cell(dbHelperRowCount, 7).Value);
                    if (string.IsNullOrEmpty((string)worksheet.Cell(dbHelperRowCount, 7).Value))
                    {
                        worksheet.Cell(dbHelperRowCount, 6).Value = "Success";
                    }
                    else
                    {
                        worksheet.Cell(dbHelperRowCount, 6).Value = "Failed";
                    }
                }

                dbHelperRowCount++;
                workbook.Save();
            }
        }
    }
}
