using NF.Tool.ExcelFlow.Common.Model.Intermediary;
using SQLite;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.Common.ExcelToDB;

public static class ExcelToSqliteBaker
{
    public static async Task Bake(ClassAndRows[] updateList, string outputPath)
    {
        using (SQLiteConnection conn = new SQLiteConnection(outputPath, SQLiteOpenFlags.ReadWrite | SQLiteOpenFlags.Create, storeDateTimeAsTicks: true))
        {
            foreach (ClassAndRows h in updateList)
            {
                Type type = h.Class;
                conn.DropTable(conn.GetMapping(type));
                conn.CreateTable(type);
            }

            conn.RunInTransaction(() =>
            {
                foreach (ClassAndRows h in updateList)
                {
                    Type type = h.Class;
                    List<object> rows = h.Rows;
                    conn.InsertAll(rows, type, runInTransaction: false);
                }
            });
        }
    }
}

