using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using dotNS;
using dotNS.Classes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NSCharts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static DotNS api = new DotNS();

        public static List<Country> countries = new List<Country>();

        public static Census mainCensus;
        public static long fromTS = -1;
        public static long toTS = -1;

        private void Form1_Load(object sender, EventArgs e)
        {
            api.UserAgent = "NSCharts dotNS library - nk.ax";
            var enumList = Enum.GetNames(typeof(Census)).ToList();
            foreach (string census in enumList)
            {
                comboBox1.Items.Add(census);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            try
            {
                var censusList = api.GetCensus(name, mainCensus, fromTS, toTS);
                countries.Add(new Country()
                {
                    name = name,
                    nodes = censusList
                });
                listBox1.Items.Add(name);
            } catch (Exception ex)
            {
                MessageBox.Show("Unknown error! The defined nation probably doesn't exist.");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null) return;
            comboBox1.Enabled = false;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = false;

            if (checkBox1.Checked)
            {
                fromTS = ((DateTimeOffset)dateTimePicker1.Value).ToUnixTimeSeconds();
                toTS = ((DateTimeOffset)dateTimePicker2.Value).ToUnixTimeSeconds();
            }
            checkBox1.Enabled = false;
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;

            mainCensus = (Census)Enum.Parse(typeof(Census), comboBox1.SelectedItem.ToString());
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                string name = listBox1.SelectedItem.ToString();
                countries.RemoveAll(z => z.name == name);
                listBox1.Items.Clear();
                countries.ForEach(z => listBox1.Items.Add(z.name));
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (countries.Count < 1) { MessageBox.Show("You didn't add a single country!"); return; }
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel files (*.xlsx)|*.xlsx";
            var drs = sfd.ShowDialog();
            if (drs != DialogResult.OK) return;
            BuildExcel(countries, sfd.FileName);
            MessageBox.Show($"Successfully saved to {sfd.FileName}!");
        }

        public static void BuildExcel(List<Country> countries, string filename)
        {
            XSSFWorkbook hssfwb = new XSSFWorkbook();
            var sheet = hssfwb.CreateSheet("Data");
            List<IRow> rows = new List<IRow>();
            for (int i = 0; i < countries.Count + 1; i++)
            {
                rows.Add(sheet.CreateRow(i));
            }
            for (int i = 0; i < countries.Count; i++)
            {
                rows[i + 1].CreateCell(0).SetCellValue(countries[i].name);
            }
            rows[0].CreateCell(0).SetCellValue("Country name");
            long lowestTimestamp = long.MaxValue;
            long highestTimestamp = long.MinValue;
            countries.ForEach(cn => { cn.nodes.ForEach(tn => { if (lowestTimestamp > tn.timestamp) lowestTimestamp = tn.timestamp; if (highestTimestamp < tn.timestamp) highestTimestamp = tn.timestamp; }); });
            var dto1 = DateTimeOffset.FromUnixTimeSeconds(lowestTimestamp);
            var dto2 = DateTimeOffset.FromUnixTimeSeconds(highestTimestamp);
            var dto3 = dto2 - dto1;
            var days = Math.Ceiling(dto3.TotalDays);
            for (int i = 0; i < days; i++)
            {
                var newDTO = dto1 + new TimeSpan(i, 0, 0, 0);
                var formed = newDTO.ToString("dd MMMM yyyy");
                rows[0].CreateCell(1 + i).SetCellValue(formed);

                countries.ForEach(cn =>
                {
                    cn.nodes.ForEach(tn =>
                    {
                        if (Math.Abs(tn.timestamp - newDTO.ToUnixTimeSeconds()) < 3600 * 6)
                        {
                            rows[countries.IndexOf(cn) + 1].CreateCell(1 + i).SetCellValue(tn.value);
                        }
                    });
                });

            }
            using (FileStream stream = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                hssfwb.Write(stream);
            }
        }
    }

    public class Country
    {
        public string name;
        public List<CensusNode> nodes;
    }
}
