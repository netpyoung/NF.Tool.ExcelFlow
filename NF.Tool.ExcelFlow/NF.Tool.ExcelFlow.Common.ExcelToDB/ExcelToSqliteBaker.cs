using NF.Tool.ExcelFlow.Common.Model.Intermediary;
using SQLite;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;

namespace NF.Tool.ExcelFlow.Common.ExcelToDB;

public static class ExcelToSqliteBaker
{
    public static async Task Bake(ClassAndRows[] updateList, string outputPath)
    {
        ClearSQLiteMappings();

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

    private static void ClearSQLiteMappings()
    {
        Type connectionType = typeof(SQLiteConnection);
        FieldInfo mappingsField = connectionType.GetField("_mappings", BindingFlags.Static | BindingFlags.NonPublic)!;
        if (mappingsField == null)
        {
            return;
        }

        Dictionary<string, TableMapping>? currentMappings = mappingsField.GetValue(null) as Dictionary<string, TableMapping>;
        if (currentMappings == null)
        {
            return;
        }
        currentMappings.Clear();
    }
}
