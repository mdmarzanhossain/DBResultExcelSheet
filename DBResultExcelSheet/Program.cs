using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DBResultExcelSheet
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string inputFilePath = @"C:\Users\Marzan Hossain\Desktop\WatcherFolder\Input.xlsx";
            string[] headerName = new string[] { "@course_code", "@course_name", "@level" , "@start_date" , "@is_active"};
            
            var workbook = new XLWorkbook(inputFilePath);
            var ws = workbook.Worksheet(1);

            int totalRow = ws.LastRowUsed().RowNumber();
            int totalColumn = ws.LastColumnUsed().ColumnNumber();

            DBHelper dBHelper = new DBHelper();
            SqlParameter[] parameters = new SqlParameter[5];
            //var value = "";

            for (int i = 2; i <= totalRow; i++)
            {
                int parameterIndex = 0;
                for (int j = 1; j <= totalColumn; j++)
                {
                    //Console.Write(ws.Cell(i, j).Value + " ");
                    parameters[parameterIndex] = new SqlParameter(headerName[parameterIndex], ws.Cell(i, j).Value);
                    parameterIndex++;
                }
                //Console.WriteLine();
                dBHelper.StoreProcedureQuery("[dbo].[usp_insert_course]", parameters[0], parameters[1], parameters[2], parameters[3], parameters[4]);
            }


            /*DBHelper dBHelper = new DBHelper();

            SqlParameter[] parameters = new SqlParameter[5];
            parameters[0] = new SqlParameter("@course_code", "");
            parameters[1] = new SqlParameter("@course_name", "Test");
            parameters[2] = new SqlParameter("@level", 500);
            parameters[3] = new SqlParameter("@start_date", DateTime.Now);
            parameters[4] = new SqlParameter("@is_active", true);

            SqlParameter[] parameters2 = new SqlParameter[5];
            parameters2[0] = new SqlParameter("@course_code", "Hello");
            parameters2[1] = new SqlParameter("@course_name", "Test");
            parameters2[2] = new SqlParameter("@level", 500);
            parameters2[3] = new SqlParameter("@start_date", DateTime.Now);
            parameters2[4] = new SqlParameter("@is_active", true);

            SqlParameter[] parameters3 = new SqlParameter[5];
            parameters3[0] = new SqlParameter("@course_code", "Hello2");
            parameters3[1] = new SqlParameter("@course_name", "Test2");
            parameters3[2] = new SqlParameter("@level", 500);
            parameters3[3] = new SqlParameter("@start_date", DateTime.Now);
            parameters3[4] = new SqlParameter("@is_active", true);


            dBHelper.StoreProcedureQuery("[dbo].[usp_insert_course]", parameters[0], parameters[1], parameters[2], parameters[3], parameters[4]);
            dBHelper.StoreProcedureQuery("[dbo].[usp_insert_course]", parameters2[0], parameters2[1], parameters2[2], parameters2[3], parameters2[4]);
            dBHelper.StoreProcedureQuery("[dbo].[usp_insert_course]", parameters3[0], parameters3[1], parameters3[2], parameters3[3], parameters3[4]);*/

            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    //services.AddHostedService<Worker>();
                });
    }
}
