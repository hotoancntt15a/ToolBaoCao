using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace ToolBaoCao
{
    public static class SQLiteCopy
    {
        private static void Main(string[] args)
        {
            string sourceConnectionString = "Data Source=source.db;";
            string targetConnectionString = "Data Source=target.db;";

            using (var sourceConnection = new SQLiteConnection(sourceConnectionString))
            using (var targetConnection = new SQLiteConnection(targetConnectionString))
            {
                sourceConnection.Open();
                targetConnection.Open();

                var tables = GetTables(sourceConnection);

                foreach (var table in tables)
                {
                    CreateTargetTable(targetConnection, table);
                    CopyData(sourceConnection, targetConnection, table);
                }
            }
        }

        private static string[] GetTables(SQLiteConnection connection)
        {
            using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", connection))
            using (var reader = command.ExecuteReader())
            {
                var tables = new System.Collections.Generic.List<string>();
                while (reader.Read())
                {
                    tables.Add(reader.GetString(0));
                }
                return tables.ToArray();
            }
        }

        private static void CreateTargetTable(SQLiteConnection targetConnection, string table)
        {
            using (var command = new SQLiteCommand($"CREATE TABLE {table} AS SELECT * FROM {table} WHERE 0;", targetConnection))
            {
                command.ExecuteNonQuery();
            }

            // Add primary key if needed
            using (var command = new SQLiteCommand($"PRAGMA table_info({table});", targetConnection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (reader["name"].ToString() == "id" && reader["pk"].ToString() == "0")
                    {
                        using (var alterCommand = new SQLiteCommand($"ALTER TABLE {table} ADD PRIMARY KEY(id);", targetConnection))
                        {
                            alterCommand.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private static void CopyData(SQLiteConnection sourceConnection, SQLiteConnection targetConnection, string table)
        {
            using (var command = new SQLiteCommand($"SELECT * FROM {table};", sourceConnection))
            using (var reader = command.ExecuteReader())
            {
                var columns = new System.Collections.Generic.List<string>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    columns.Add(reader.GetName(i));
                }

                var batchSize = 1000;
                var batch = new System.Collections.Generic.List<string[]>();
                var insertCommand = BuildInsertCommand(table, columns);

                while (reader.Read())
                {
                    var values = new string[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        values[i] = reader[i].ToString();
                        if (columns[i] == "ho_ten" || columns[i] == "ngay_sinh")
                        {
                            values[i] = ComputeMD5(values[i]);
                        }
                    }
                    batch.Add(values);

                    if (batch.Count >= batchSize)
                    {
                        ExecuteBatchInsert(targetConnection, insertCommand, batch);
                        batch.Clear();
                    }
                }

                if (batch.Count > 0)
                {
                    ExecuteBatchInsert(targetConnection, insertCommand, batch);
                }
            }
        }

        private static string BuildInsertCommand(string table, System.Collections.Generic.List<string> columns)
        {
            var columnList = string.Join(", ", columns);
            var paramList = string.Join(", ", columns.ConvertAll(col => "@" + col));
            return $"INSERT INTO {table} ({columnList}) VALUES ({paramList});";
        }

        private static void ExecuteBatchInsert(SQLiteConnection connection, string commandText, System.Collections.Generic.List<string[]> batch)
        {
            using (var transaction = connection.BeginTransaction())
            using (var command = new SQLiteCommand(commandText, connection))
            {
                foreach (var values in batch)
                {
                    command.Parameters.Clear();
                    for (int i = 0; i < values.Length; i++)
                    {
                        command.Parameters.AddWithValue("@" + command.Parameters[i].ParameterName, values[i]);
                    }
                    command.ExecuteNonQuery();
                }
                transaction.Commit();
            }
        }

        private static string ComputeMD5(string input)
        {
            using (var md5 = MD5.Create())
            {
                var inputBytes = Encoding.UTF8.GetBytes(input);
                var hashBytes = md5.ComputeHash(inputBytes);
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
            }
        }
    }
}