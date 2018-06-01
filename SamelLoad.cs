using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;

namespace ReportGenerations
{
    class SamelLoad
    {
        public void startGenerate(DateTime dateFrom, DateTime dateTo, DataBase db, bool NewFormat)
        {

            var DST = new DataSetTable();
            try
            {
                DST.values.Add("month", BaseClass.getMonthName(dateTo.Month) + " " + dateTo.ToString("yyyy"));
                DST.values.Add("dateTo", dateTo.ToString("dd/MM/yyyy"));
                DST.values.Add("year", dateTo.Year.ToString());
                DateTime dateFromYear = new DateTime(dateTo.Year, 1, 1);

                Dictionary<string, string> destinationPathes = new Dictionary<string, string>();
                Dictionary<string, string> templatePaths = new Dictionary<string, string>();
                string destinationPath = @"c:\dest";
                string templatePath = @"c:\templ"; ;
                destinationPathes.Add("RegistersRemains", destinationPath + $"RegistersRemains_{dateTo.ToString("yyyyMMdd")}.xlsx");
                //destinationPathes.Add("provision", destinationPath + $"provision_{dateTo.ToString("yyyyMMdd")}.xlsx");
                //destinationPathes.Add("DealsReport", destinationPath + $"DealsReport_{dateTo.ToString("yyyyMMdd")}.xlsx");
                foreach (string typeList in destinationPathes.Keys)
                {
                    loadData(DST, db, dateFrom, dateTo, typeList);
                    ReportsGeneration.ReportGeneration(templatePath + typeList + ".xlsx", DST, destinationPathes[typeList]);
                }

            }
            }
        private void loadData(DataSetTable dST, DataBase db, DateTime dateFrom, DateTime dateTo, string typelist)
        {

            try
            {
                //  string dir = AppDomain.CurrentDomain.BaseDirectory + @"SQLFiles\ЗапросыОтчеБугалтерия";
                string dir = AppDomain.CurrentDomain.BaseDirectory + @"SQLFiles\ОтчетДляРуководства";
                var sqlFiles = System.IO.Directory.GetFiles(dir, typelist + "$*");
                DataTable dt = new DataTable();
                foreach (var file in sqlFiles)
                {

                    var nameTable = file.Substring(file.IndexOf("$") + 1);
                    if (nameTable.Contains("$"))
                        nameTable = nameTable.Substring(0, nameTable.IndexOf("$"));
                    if (!dST.Tables.Contains(nameTable))
                    {
                        dt = new DataTable(nameTable);
                        dST.Tables.Add(dt);
                    }
                    else dt.PrimaryKey = new DataColumn[1] { dt.Columns[0] };
                    string requestStr = SQL2str.translateSQL2str(file, dateFrom, dateTo);
                    try
                    {
                        //добавим таблицу в набор
                        using (OracleDataReader reader = db.executeQuery(requestStr))
                            dST.Load(reader, LoadOption.OverwriteChanges, dt);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Ошибка выгрузки тиблиы '" + nameTable + "' :" + ex.Message, ex);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new DbException("При генерации данных для отчета для руководства (вкладка \"Динамика остатков за месяц\") возникла ошибка.", ex);
            }

        }
    }
}
