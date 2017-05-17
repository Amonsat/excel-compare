using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;
using ExcelCompare.Properties;
using Timer = System.Windows.Forms.Timer;

namespace ExcelCompare {
    public partial class Form1 : Form {
        private ExcelFile _file1;
        private ExcelFile _file2;
        private ExcelFile.Sheet _sheet1;
        private ExcelFile.Sheet _sheet2;
        private ExcelFile.Sheet _comparedSheet;
        private readonly Compare _compare = new Compare();
        private TimeSpan _time;
        private Timer _timer = new Timer(){Interval = 1000};
        private readonly BackgroundWorker _worker = new BackgroundWorker();
        private readonly BackgroundWorker _saveWorker = new BackgroundWorker();
        private ExcelFile _comparedFile = new ExcelFile();


        public Form1() {
            InitializeComponent();
            _timer.Tick += delegate {
                _time += new TimeSpan(0, 0, 1);
                lTime.Text = _time.ToString();
            };
            _worker.DoWork += Worker_DoWork;
            _worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            _saveWorker.DoWork += SaveWorker_DoWork;
            _saveWorker.RunWorkerCompleted += SaveWorker_RunWorkerCompleted;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e) {
            _comparedSheet = _compare.CompareSheets(ref _sheet1, ref _sheet2);
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            _timer.Stop();
            var saveFileDialog = new SaveFileDialog {
                Filter = Resources.ExcelFilesMask,
                FileName = "Compared.xls"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                _timer.Start();
                lStep.Text = "Шаг 3: Сохраняем";
                _comparedFile.Name = saveFileDialog.FileName;
                _comparedFile.Sheets.Add(_comparedSheet);    
                _saveWorker.RunWorkerAsync();
            }
        }

        private void SaveWorker_DoWork(object sender, DoWorkEventArgs e) {
            _comparedFile.Save();
        }

        private void SaveWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            _timer.Stop();
            lStep.Text = "Выполнено";
            pbProgress.Value = 0;
        }

        private void bFile1Open_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog {
                Filter = Resources.ExcelFilesMask
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                tbFile1Path.Text = openFileDialog.FileName;
                _file1 = new ExcelFile(tbFile1Path.Text);
                cbFile1SheetList.Items.Clear();
                foreach (var sheet in _file1.Sheets) {
                    cbFile1SheetList.Items.Add(sheet.Name);
                }
                if (cbFile1SheetList.Items.Count > 0) cbFile1SheetList.SelectedIndex = 0;
            }

        }

        private void cbFile1SheetList_SelectedIndexChanged(object sender, EventArgs e) {
            cbFile1ColumnList.Items.Clear();
            foreach (var column in _file1.Sheets[cbFile1SheetList.SelectedIndex].Data.Columns) {
                cbFile1ColumnList.Items.Add(column.ToString());
            }
            if (cbFile1ColumnList.Items.Count > 0) cbFile1ColumnList.SelectedIndex = 0;

        }

        private void cbFile1ColumnList_SelectedIndexChanged(object sender, EventArgs e) {
            _file1.Sheets[cbFile1SheetList.SelectedIndex].KeyColumnIndex = cbFile1ColumnList.SelectedIndex;
        }

        private void bFile2Open_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog {
                Filter = Resources.ExcelFilesMask
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                tbFile2Path.Text = openFileDialog.FileName;
                _file2 = new ExcelFile(tbFile2Path.Text);
                cbFile2SheetList.Items.Clear();
                foreach (var sheet in _file2.Sheets) {
                    cbFile2SheetList.Items.Add(sheet.Name);
                }
                if (cbFile2SheetList.Items.Count > 0) cbFile2SheetList.SelectedIndex = 0;
            }
        }

        private void cbFile2SheetList_SelectedIndexChanged(object sender, EventArgs e) {
            cbFile2ColumnList.Items.Clear();
            foreach (var column in _file2.Sheets[cbFile2SheetList.SelectedIndex].Data.Columns) {
                cbFile2ColumnList.Items.Add(column.ToString());
            }
            if (cbFile2ColumnList.Items.Count > 0) cbFile2ColumnList.SelectedIndex = 0;

        }

        private void cbFile2ColumnList_SelectedIndexChanged(object sender, EventArgs e) {
            _file2.Sheets[cbFile2SheetList.SelectedIndex].KeyColumnIndex = cbFile2ColumnList.SelectedIndex;
        }

        private void bCompare_Click(object sender, EventArgs e) {
            _compare.Progress += ProgressProcessor;
            _sheet1 = _file1.Sheets[cbFile1SheetList.SelectedIndex];
            _sheet2 = _file2.Sheets[cbFile2SheetList.SelectedIndex];
            _time = new TimeSpan();
            _timer.Start();
            _worker.RunWorkerAsync();
        }
        private void ProgressProcessor(int current, int max, string step) {
            if (!pbProgress.InvokeRequired) {
                pbProgress.Maximum = max;
                pbProgress.Value = current;
            } else pbProgress.Invoke(new Compare.EventProgress(ProgressProcessor), current, max, step);
            if (!lStep.InvokeRequired) {
                lStep.Text = step;
            } else lStep.Invoke(new Compare.EventProgress(ProgressProcessor), current, max, step);
        }

    }
}
