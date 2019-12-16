using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using TikaOnDotNet.TextExtraction;

namespace EmailFromPDF
{
    public partial class Form1 : Form
    {
        private string[] _emails;
        private string _putanjaDest;
        private static int _fileNo;
        private string[] _files;

        public Form1()
        {
            _files = new string[20];
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            var fdlg = new OpenFileDialog
            {
                Title = "Izaberite pdf fajl za pretragu",
                InitialDirectory = @"c:\",
                Filter = "Pdf Files|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true
            };
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                _files[_fileNo] = fdlg.FileName;
                lblFiles.Rows.Add();
                lblFiles.Rows[_fileNo].Cells["Path"].Value = _files[_fileNo];
                lblFiles.Rows[_fileNo].Cells["Files"].Value = _files[_fileNo].Substring(_files[_fileNo].LastIndexOf("\\", StringComparison.Ordinal) + 1);
                _fileNo++;
                btnExport.Enabled = true;
                btnRemove.Enabled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Loader(1);
            _emails = ProcitajEmailove();
            DataGridViewPopuni();
            Loader(0);


        }

        private void DataGridViewPopuni()
        {
            dataGridView1.DataSource = _emails.Select((x, index) =>
                new { No = index + 1, Email = x }).OrderBy(x => x.No).ToList();

            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].Width = 30;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

        }

        private string[] ProcitajEmailove()
        {
            var text = "";
            foreach (var file in _files)
            {
                if(file != null)
                text += new TextExtractor().Extract(file).Text;
            }

            const string MatchEmailPattern =
                @"(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@"
                + @"((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
                + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|"
                + @"([a-zA-Z]+[\w-]+\.)+[a-zA-Z]{2,4})";
            Regex rx = new Regex(MatchEmailPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);
            MatchCollection matches = rx.Matches(text);
            int noOfMatches = matches.Count;
            var tekst = new string[noOfMatches];
            for (int i = 0; i < matches.Count; i++)
            {
                tekst[i] = matches[i].Value;

            }
            return tekst.Distinct().ToArray();
        }

        private void Loader(int i)
        {
            if (i == 1)
            {
                var w = new Form() { Size = new Size(0, 0) };
                Task.Delay(TimeSpan.FromSeconds(1))
                    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());

                MessageBox.Show(w, "Fajl se obradjuje, sacekajte...", "Obrada");
                button1.Enabled = false;
                btnSave.Enabled = false;
                btnExport.Enabled = false;
            }
            else
            {
                MessageBox.Show("Zavrseno!!!");
                button1.Enabled = true;
                btnSave.Enabled = true;
                btnExport.Enabled = true;
                panelSave.Enabled = true;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            var folderBroswer = new FolderBrowserDialog
            {
                RootFolder = Environment.SpecialFolder.MyComputer
            };

            if (folderBroswer.ShowDialog() == DialogResult.OK)
            {
                _putanjaDest = folderBroswer.SelectedPath;
            }

            if (radioExcel.Checked)
            {
                new SaveInExcel(_emails, _putanjaDest);
            }
            else
            {
                new SaveInTxt(_emails, _putanjaDest);
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var frm = new About();
            frm.ShowDialog();
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            var selectedIndex = lblFiles.SelectedCells[0].RowIndex;
            var selektovanFile = Convert.ToString(lblFiles.Rows[selectedIndex].Cells["Path"].Value);
            List<string> filesList = new List<string>(_files);
            filesList.Remove(selektovanFile);
            filesList.Add(null);
            _files = filesList.ToArray();
            lblFiles.Rows.Remove(lblFiles.Rows[selectedIndex]);
            _fileNo--;
            if (_fileNo < 1)
            {
                btnExport.Enabled = false;
                btnRemove.Enabled = false;
            }
        }

    }
}
