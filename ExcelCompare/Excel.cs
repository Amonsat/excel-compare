using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

    namespace Excel {
        internal class ExcelFile {
            public string Name { get; set; }
            public List<Sheet> Sheets { get; set; }

            public ExcelFile() {
                Sheets = new List<Sheet>();
            }

            public ExcelFile(string filename) {
                Name = filename;
                ReadExcelFile(Name);
            }

            private void ReadExcelFile(string fileName) {
                var connectionString = GetConnectionString(fileName);
                using (var conn = new OleDbConnection(connectionString)) {
                    conn.Open();
                    var cmd = new OleDbCommand {Connection = conn};
                    var dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dtSheet != null) {
                        Sheets = new List<Sheet>();
                        foreach (
                            var sheetName in
                                dtSheet.Rows.Cast<DataRow>()
                                    .Select(dr => dr["TABLE_NAME"].ToString())
                                    .Where(sheetName => sheetName.EndsWith("$"))) {
                            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                            var sheet = new Sheet {Name = sheetName};
                            var da = new OleDbDataAdapter(cmd);
                            da.Fill(sheet.Data);
                            Sheets.Add(sheet);
                        }
                    }
                    conn.Close();
                }
            }

            private static string GetConnectionString(string fileName) {
                var props = new Dictionary<string, string>();

                // XLSX - Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
                props["Data Source"] = fileName;

                // XLS - Excel 2003 and Older
                //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                //props["Extended Properties"] = "Excel 8.0";
                //props["Data Source"] = "C:\\MyExcel.xls";

                var sb = new StringBuilder();

                foreach (var prop in props) {
                    sb.Append(prop.Key);
                    sb.Append('=');
                    sb.Append(prop.Value);
                    sb.Append(';');
                }

                return sb.ToString();
            }

            public void Save() {
                WriteExcelFile(Name);
            }
            private void WriteExcelFile(string fileName) {
                var connectionString = GetConnectionString(fileName);
                using (var conn = new OleDbConnection(connectionString)) {
                    conn.Open();

                    var columns = String.Empty;
                    var columnsWithType = String.Empty;
                    foreach (var column in Sheets[0].Data.Columns) {
                        columns += String.Format("'{0}', ", column.ToString());
                        columnsWithType += String.Format("'{0}' VARCHAR, ", column.ToString());
                    }
                    columns = columns.Trim().TrimEnd(',');
                    columnsWithType = columnsWithType.Trim().TrimEnd(',');
                    
                    var cmd = new OleDbCommand {
                        Connection = conn,
                        CommandText = String.Format("CREATE TABLE [Compared] ({0});", columnsWithType)
                    };
                    cmd.ExecuteNonQuery();
                    foreach (DataRow row in Sheets[0].Data.Rows) {
//                        for (var i = 0; i < Sheets[0].Data.Columns.Count; i++) {
                            cmd.CommandText = String.Format("INSERT INTO [compared]({0}) VALUES('{1}'); ", columns, String.Join("', '", row.ItemArray));
                            cmd.ExecuteNonQuery();
//                        }                        
                    }
                    conn.Close();
                }
            }

            internal class Sheet {
                public int KeyColumnIndex { get; set; }
                public DataTable Data { get; set; }
                private string _name;

                public Sheet() {
                    Data = new DataTable();
                }
                public string Name {
                    get { return _name; }
                    set {
                        _name = value.TrimEnd('$');
                        Data.TableName = _name;
                    }
                }
                
            }

        }

        internal class Compare {
            public List<string> File1Style;
            public List<string> File2Style;

            private string _step;
            private int _max;
            private int _current;

            public delegate void EventProgress(int current, int max, string step);

            public event EventProgress Progress;

            private void SetStep(string step, int max) {
                _step = step;
                _max = max;
                _current = 0;
            }

            private void ProgressUp() {
                Progress(_current++, _max, _step);
            }

            private static void SortByKeyColumn(ExcelFile.Sheet sheet) {
                sheet.Data.DefaultView.Sort = sheet.Data.Columns[sheet.KeyColumnIndex].ColumnName + " ASC";
            }

            private void TrimKeyColumn(ExcelFile.Sheet sheet) {
                foreach (DataRow row in sheet.Data.Rows) {
                    row[sheet.KeyColumnIndex] = row[sheet.KeyColumnIndex].ToString().Trim();
                    ProgressUp();
                }
            }


            public ExcelFile.Sheet CompareSheets(ref ExcelFile.Sheet sheet1, ref ExcelFile.Sheet sheet2) {
//            public void CompareSheets(ref ExcelFile.Sheet sheet1, ref ExcelFile.Sheet sheet2) {
                SetStep("Шаг 1: Анализируем", sheet1.Data.Rows.Count + sheet2.Data.Rows.Count);
                SortByKeyColumn(sheet1);
                TrimKeyColumn(sheet1);
                SortByKeyColumn(sheet2);
                TrimKeyColumn(sheet2);

                






                var dt1Unique = sheet1.Data.DefaultView.ToTable(true, sheet1.Data.Columns[sheet1.KeyColumnIndex].ColumnName);
                var dt2Unique = sheet2.Data.DefaultView.ToTable(true, sheet2.Data.Columns[sheet2.KeyColumnIndex].ColumnName);
                dt1Unique.Columns[sheet1.KeyColumnIndex].ColumnName = "keyCol";
                dt2Unique.Columns[sheet2.KeyColumnIndex].ColumnName = "keyCol";
                dt1Unique.Merge(dt2Unique);                
//                ExcelFile.Sheet dtUnique.Data = dt1Unique.DefaultView.ToTable(true, "key");
                var dtUnique = new ExcelFile.Sheet {
                    Name = "Compared",
                    Data = dt1Unique.DefaultView.ToTable(true, "keyCol")
                };
                dtUnique.Data.DefaultView.Sort = "keyCol" + " ASC";
                SetStep("Шаг 2: Обрабатываем", dtUnique.Data.Rows.Count + sheet1.Data.Columns.Count + sheet2.Data.Columns.Count);
//                ProgressUp();
                var comapredSheet1 = new ExcelFile.Sheet();
                var comparedSheet2 = new ExcelFile.Sheet();
                foreach (DataColumn column in sheet1.Data.Columns) {
                    comapredSheet1.Data.Columns.Add(column.ColumnName);
                    ProgressUp();
                }
                foreach (DataColumn column in sheet2.Data.Columns) {
                    comparedSheet2.Data.Columns.Add(column.ColumnName);
                    ProgressUp();
                }
                foreach (DataRow row in dtUnique.Data.Rows) {
                    var dt1FindString = sheet1.Data.Columns[sheet1.KeyColumnIndex].ColumnName + " Like '" +
                                        row[sheet1.KeyColumnIndex].ToString() + "'";
                    foreach (var dt1Row in sheet1.Data.Select(dt1FindString)) {
                        comapredSheet1.Data.Rows.Add(dt1Row.ItemArray);
                    }
                    var dt2FindString = sheet2.Data.Columns[sheet2.KeyColumnIndex].ColumnName + " Like '" +
                                        row[sheet2.KeyColumnIndex].ToString() + "'";
                    foreach (var dt2Row in sheet2.Data.Select(dt2FindString)) {
                        comparedSheet2.Data.Rows.Add(dt2Row.ItemArray);
                    }
                    var dt = comapredSheet1.Data.Rows.Count < comparedSheet2.Data.Rows.Count ? comapredSheet1 : comparedSheet2;
                    for (var i = 0; i < Math.Abs(comapredSheet1.Data.Rows.Count - comparedSheet2.Data.Rows.Count); i++) {
                        dt.Data.Rows.Add(String.Empty);
                    }
                    ProgressUp();
                }
                sheet1 = comapredSheet1;
                sheet2 = comparedSheet2;
                
                
//                dtUnique.Data = sheet1.Data;
                dtUnique.Data.Columns.Add(new DataColumn("empty"));
                for (var i = 0; i < sheet1.Data.Columns.Count; i++) {
                    dtUnique.Data.Columns.Add(sheet1.Data.Columns[i] + "_sh1");
                    for (var j = 0; j < sheet1.Data.Rows.Count; j++) {
                        dtUnique.Data.Rows[j][1 + i] = sheet1.Data.Rows[j][i];
                    }
                }
                dtUnique.Data.Columns.Add(new DataColumn("empty"));
                for (var i = 0; i < sheet2.Data.Columns.Count; i++) {
                    dtUnique.Data.Columns.Add(sheet2.Data.Columns[i]+"_sh2");
                    for (var j = 0; j < sheet2.Data.Rows.Count; j++) {
                        dtUnique.Data.Rows[j][10 + i] = sheet2.Data.Rows[j][i];
                    }
                }
                

                return dtUnique;









                // emplicit ixplicit

//                SetStep("Шаг 2: Объединение", sheet1.Columns.Count + sheet2.Columns.Count + dtUnique.Rows.Count);

//                var comapredSheet1 = new ExcelFile.Sheet();
//                var comparedSheet2 = new ExcelFile.Sheet();
//                foreach (DataColumn column in sheet1.Columns) {
//                    comapredSheet1.Columns.Add(column.ColumnName);
//                    ProgressUp();
//                }
//                foreach (DataColumn column in sheet2.Columns) {
//                    comparedSheet2.Columns.Add(column.ColumnName);
//                    ProgressUp();
//                }
//                foreach (DataRow row in dtUnique.Rows) {
//                    var dt1FindString = sheet1.Columns[sheet1.KeyColumnIndex].ColumnName + " Like '" +
//                                        row[sheet1.KeyColumnIndex].ToString() + "'";
//                    foreach (var dt1Row in sheet1.Select(dt1FindString)) {
//                        comapredSheet1.Rows.Add(dt1Row.ItemArray);
//                    }
//                    var dt2FindString = sheet2.Columns[sheet2.KeyColumnIndex].ColumnName + " Like '" +
//                                        row[sheet2.KeyColumnIndex].ToString() + "'";
//                    foreach (var dt2Row in sheet2.Select(dt2FindString)) {
//                        comparedSheet2.Rows.Add(dt2Row.ItemArray);
//                    }
//                    var dt = comapredSheet1.Rows.Count < comparedSheet2.Rows.Count ? comapredSheet1 : comparedSheet2;
//                    for (var i = 0; i < Math.Abs(comapredSheet1.Rows.Count - comparedSheet2.Rows.Count); i++) {
//                        dt.Rows.Add(String.Empty);
//                    }
//                    ProgressUp();
//                }
//                sheet1 = comapredSheet1;
//                sheet2 = comparedSheet2;


//                File1Style = new List<string>();
//                File2Style = new List<string>();
//                SetStep("Шаг 3", sheet1.Columns.Count * sheet1.Rows.Count);
//
//                for (var row = 0; row < sheet1.Rows.Count - 1; row++) {
//                    for (var col = 0; col < sheet1.Columns.Count - 1; col++) {
//                        var cell1 = sheet1.Rows[row][col];
//                        var cell2 = sheet2.Rows[row][col];
//                        if (!cell1.Equals(cell2)) {
//                            File1Style.Add(row.ToString() + ';' + col.ToString());
//                            File2Style.Add(row.ToString() + ';' + col.ToString());
//                        }
//                    }
//                    ProgressUp();
//                }
            }
        }

    }