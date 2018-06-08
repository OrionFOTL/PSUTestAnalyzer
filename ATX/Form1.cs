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

namespace ATX
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public DataTable dataTabelaLoad = new DataTable();
        public DataTable dataTabelaRipple = new DataTable();
        public bool fileLoaded = false;

        public void CSVdoDataTable(string filePath)
        {
            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] naglowki = sr.ReadLine().Split(',');
                dataTabelaRipple.Columns.Add(naglowki[0]);
                for (int i = 0; i < naglowki.Length; i++)
                {
                    if (i < naglowki.Length-3)
                    {
                        dataTabelaLoad.Columns.Add(naglowki[i]);
                    }
                    else
                    {
                        dataTabelaRipple.Columns.Add(naglowki[i]);
                    }
                }
                while (!sr.EndOfStream)
                {
                    string[] wiersze = sr.ReadLine().Split(',');
                    DataRow dataWierszLoad = dataTabelaLoad.NewRow();
                    DataRow dataWierszRipple = dataTabelaRipple.NewRow();
                    int n = 0;
                    dataWierszRipple[n] = Convert.ToDouble(wiersze[0].Replace('.', ','));
                    for (int i = 0; i < naglowki.Length; i++)
                    {
                        if (i < naglowki.Length-3)
                        {
                            dataWierszLoad[i] = Convert.ToDouble(wiersze[i].Replace('.', ','));
                        }
                        else
                        {
                            dataWierszRipple[++n] = Convert.ToDouble(wiersze[i].Replace('.', ','));
                        }
                        
                    }
                    dataTabelaLoad.Rows.Add(dataWierszLoad);
                    dataTabelaRipple.Rows.Add(dataWierszRipple);
                }
            }
        }
        public void wypelnijLoadTable(DataTable dataTabela)
        {
            dataGridViewLoad.DataSource = dataTabela;
            dataGridViewLoad.CellFormatting += 
                new DataGridViewCellFormattingEventHandler(dataGridViewLoad_CellFormatting);
            formatujLoadTable();
        }
        public void wypelnijRippleTable(DataTable dataTabela)
        {
            dataGridViewRipple.DataSource = dataTabela;
            dataGridViewRipple.CellFormatting +=
                new DataGridViewCellFormattingEventHandler(dataGridViewRipple_CellFormatting);
            formatujRippleTable();
        }
        public void formatujLoadTable()
        {
            foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
            {
                if (Convert.ToDouble(wiersz.Cells[1].Value) < 11.4 || Convert.ToDouble(wiersz.Cells[1].Value) > 12.6)
                {
                    wiersz.Cells[1].Style.BackColor = Color.Red;
                }
                else wiersz.Cells[1].Style.BackColor = Color.White;
                if (Convert.ToDouble(wiersz.Cells[2].Value) < 4.75 || Convert.ToDouble(wiersz.Cells[2].Value) > 5.25)
                {
                    wiersz.Cells[2].Style.BackColor = Color.Red;
                }
                else wiersz.Cells[2].Style.BackColor = Color.White;
                if (Convert.ToDouble(wiersz.Cells[3].Value) < 3.135 || Convert.ToDouble(wiersz.Cells[3].Value) > 3.465)
                {
                    wiersz.Cells[3].Style.BackColor = Color.Red;
                }
                else wiersz.Cells[3].Style.BackColor = Color.White;
            }
        }
        public void formatujRippleTable()
        {
            foreach (DataGridViewRow wiersz in dataGridViewRipple.Rows)
            {
                if (Convert.ToDouble(wiersz.Cells[1].Value.ToString().Replace('.', ',')) > 120)
                {
                    wiersz.Cells[1].Style.BackColor = Color.Red;
                }
                else wiersz.Cells[1].Style.BackColor = Color.White;
                if (Convert.ToDouble(wiersz.Cells[2].Value.ToString().Replace('.', ',')) > 50)
                {
                    wiersz.Cells[2].Style.BackColor = Color.Red;
                }
                else wiersz.Cells[2].Style.BackColor = Color.White;
                if (Convert.ToDouble(wiersz.Cells[3].Value.ToString().Replace('.', ',')) > 50)
                {
                    wiersz.Cells[3].Style.BackColor = Color.Red;
                }
                else wiersz.Cells[3].Style.BackColor = Color.White;
            }
        }
        void dataGridViewLoad_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            string[] formats = { "00%", "0.000V", "0.000V", "0.000V", "###.00W", "#.00%", "0 RPM",
                                "0.# db(A)"};
            e.Value = double.Parse(e.Value.ToString()).ToString(formats[e.ColumnIndex]);
        }
        void dataGridViewRipple_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            string[] formats = { "00%", "0.0 mV", "0.0 mV", "0.0 mV" };
            e.Value = double.Parse(e.Value.ToString()).ToString(formats[e.ColumnIndex]);
        }

        public void wypelnijVregChart()
        {
            chartVReg.Series["Series1"].Points.Clear();
            if (radioButton12V.Checked)
            {
                chartVReg.ChartAreas[0].AxisY.LabelStyle.Interval = 0.25;
                chartVReg.ChartAreas[0].AxisY.Minimum = 11.25;
                chartVReg.ChartAreas[0].AxisY.Maximum = 12.75;
                chartVReg.ChartAreas[0].AxisY.Interval = 0.25;
                chartVReg.ChartAreas[0].AxisY.MajorGrid.Interval = 5;
                chartVReg.ChartAreas[0].AxisY.MinorGrid.Interval = 0.25;
                chartVReg.ChartAreas[0].AxisY.MajorGrid.IntervalOffset = -4.25;
                chartVReg.ChartAreas[0].AxisY.MajorTickMark.Interval = 0.25;
                chartVReg.ChartAreas[0].AxisY.StripLines[0].IntervalOffset = 0.15;
                chartVReg.ChartAreas[0].AxisY.StripLines[1].IntervalOffset = 1.35;
                chartVReg.Series[0].Color = Color.Gold;

                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    chartVReg.Series["Series1"].Points.AddXY(
                    Convert.ToDouble(wiersz.Cells[0].Value.ToString()) *
                    Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[4].Value.ToString()),
                    Convert.ToDouble(wiersz.Cells[1].Value.ToString()));
                }
            }
            else if (radioButton5V.Checked)
            {
                chartVReg.ChartAreas[0].AxisY.LabelStyle.Interval = 0.1;
                chartVReg.ChartAreas[0].AxisY.Minimum = 4.7;
                chartVReg.ChartAreas[0].AxisY.Maximum = 5.3;
                chartVReg.ChartAreas[0].AxisY.Interval = 0.1;
                chartVReg.ChartAreas[0].AxisY.MajorGrid.Interval = 0.6;
                chartVReg.ChartAreas[0].AxisY.MinorGrid.Interval = 0.1;
                chartVReg.ChartAreas[0].AxisY.MajorGrid.IntervalOffset = -0.3;
                chartVReg.ChartAreas[0].AxisY.MajorTickMark.Interval = 0.1;
                chartVReg.ChartAreas[0].AxisY.StripLines[0].IntervalOffset = 0.05;
                chartVReg.ChartAreas[0].AxisY.StripLines[1].IntervalOffset = 0.55;
                chartVReg.Series[0].Color = Color.Red;

                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    chartVReg.Series["Series1"].Points.AddXY(
                    Convert.ToDouble(wiersz.Cells[0].Value.ToString()) *
                    Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[4].Value.ToString()),
                    Convert.ToDouble(wiersz.Cells[2].Value.ToString()));
                }
            }
            else if (radioButton3V.Checked)
            {
                chartVReg.ChartAreas[0].AxisY.LabelStyle.Interval = 0.1;
                chartVReg.ChartAreas[0].AxisY.Minimum = 3.1;
                chartVReg.ChartAreas[0].AxisY.Maximum = 3.5;
                chartVReg.ChartAreas[0].AxisY.Interval = 0.05;
                chartVReg.ChartAreas[0].AxisY.MajorGrid.Interval = 0.5;
                chartVReg.ChartAreas[0].AxisY.MinorGrid.Interval = 0.05;
                chartVReg.ChartAreas[0].AxisY.MajorGrid.IntervalOffset = -0.3;
                chartVReg.ChartAreas[0].AxisY.MajorTickMark.Interval = 0.1;
                chartVReg.ChartAreas[0].AxisY.StripLines[0].IntervalOffset = 0.035;
                chartVReg.ChartAreas[0].AxisY.StripLines[1].IntervalOffset = 0.365;
                chartVReg.Series[0].Color = Color.DarkOrange;

                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    chartVReg.Series["Series1"].Points.AddXY(
                    Convert.ToDouble(wiersz.Cells[0].Value.ToString()) *
                    Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[4].Value.ToString()),
                    Convert.ToDouble(wiersz.Cells[3].Value.ToString()));
                }
            }
        }
        public void wypelnijRippleChart()
        {
            chartRipple.Series["Series1"].Points.Clear();
            if (radioButton12VRipple.Checked)
            {
                chartRipple.ChartAreas[0].AxisY.LabelStyle.Interval = 0.2;
                chartRipple.ChartAreas[0].AxisY.Maximum = 140;
                chartRipple.ChartAreas[0].AxisY.LabelStyle.Interval = 20;
                chartRipple.ChartAreas[0].AxisY.MajorGrid.Interval = 20;
                chartRipple.ChartAreas[0].AxisY.MajorTickMark.Interval = 20;
                chartRipple.ChartAreas[0].AxisY.StripLines[0].IntervalOffset = 120;
                chartRipple.Series[0].Color = Color.Gold;

                foreach (DataGridViewRow wiersz in dataGridViewRipple.Rows)
                {
                    chartRipple.Series["Series1"].Points.AddXY(
                    Convert.ToDouble(wiersz.Cells[0].Value.ToString()),
                    Convert.ToDouble(wiersz.Cells[1].Value.ToString()));
                }
            }
            else if (radioButton5VRipple.Checked)
            {
                chartRipple.ChartAreas[0].AxisY.LabelStyle.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.Maximum = 60;
                chartRipple.ChartAreas[0].AxisY.LabelStyle.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.MajorGrid.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.MajorTickMark.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.StripLines[0].IntervalOffset = 50;
                chartRipple.Series[0].Color = Color.Red;

                foreach (DataGridViewRow wiersz in dataGridViewRipple.Rows)
                {
                    chartRipple.Series["Series1"].Points.AddXY(
                    Convert.ToDouble(wiersz.Cells[0].Value.ToString()),
                    Convert.ToDouble(wiersz.Cells[2].Value.ToString()));
                }
            }
            else if (radioButton3VRipple.Checked)
            {
                chartRipple.ChartAreas[0].AxisY.LabelStyle.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.Maximum = 60;
                chartRipple.ChartAreas[0].AxisY.LabelStyle.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.MajorGrid.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.MajorTickMark.Interval = 10;
                chartRipple.ChartAreas[0].AxisY.StripLines[0].IntervalOffset = 50;
                chartRipple.Series[0].Color = Color.DarkOrange;

                foreach (DataGridViewRow wiersz in dataGridViewRipple.Rows)
                {
                    chartRipple.Series["Series1"].Points.AddXY(
                    Convert.ToDouble(wiersz.Cells[0].Value.ToString()),
                    Convert.ToDouble(wiersz.Cells[3].Value.ToString()));
                }
            }
        }
        public void wypelnijNoiseChart()
        {
            chartNoise.Series["RPM"].Points.Clear();
            chartNoise.Series["dBA"].Points.Clear();

            foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
            {
                chartNoise.Series["RPM"].Points.AddXY(
                Convert.ToDouble(wiersz.Cells[0].Value.ToString()) *
                Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[4].Value.ToString()),
                Convert.ToDouble(wiersz.Cells[6].Value.ToString()));

                chartNoise.Series["dBA"].Points.AddXY(
                Convert.ToDouble(wiersz.Cells[0].Value.ToString()) *
                Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[4].Value.ToString()),
                Convert.ToDouble(wiersz.Cells[7].Value.ToString()));
            }
        }
        public void wypelnijPodsumowanie()
        {
            /////Najgorsza regulacja napięć
            {
                double currentMax = 0.0, maxIndex = 0.0;
                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    if ((Math.Abs(Convert.ToDouble(wiersz.Cells[1].Value.ToString()) - 12) / 12) > currentMax)
                    {
                        currentMax = (Math.Abs(Convert.ToDouble(wiersz.Cells[1].Value.ToString()) - 12) / 12);
                        maxIndex = Convert.ToDouble(wiersz.Cells[0].Value.ToString());
                    }
                }
                labelReg12V.Text = String.Format("12V: {0:0.00%} @ {1:00%}", currentMax, maxIndex);
                if (currentMax > 0.05) labelReg12V.ForeColor = Color.Red;

                currentMax = 0.0;
                maxIndex = 0.0;
                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    if ((Math.Abs(Convert.ToDouble(wiersz.Cells[2].Value.ToString()) - 5) / 5) > currentMax)
                    {
                        currentMax = (Math.Abs(Convert.ToDouble(wiersz.Cells[2].Value.ToString()) - 5) / 5);
                        maxIndex = Convert.ToDouble(wiersz.Cells[0].Value.ToString());
                    }
                }
                labelReg5V.Text = String.Format("5V: {0:0.00%} @ {1:00%}", currentMax, maxIndex);
                if (currentMax > 0.05) labelReg5V.ForeColor = Color.Red;

                currentMax = 0.0;
                maxIndex = 0.0;
                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    if ((Math.Abs(Convert.ToDouble(wiersz.Cells[3].Value.ToString()) - 3.3) / 3.3) > currentMax)
                    {
                        currentMax = (Math.Abs(Convert.ToDouble(wiersz.Cells[3].Value.ToString()) - 3.3) / 3.3);
                        maxIndex = Convert.ToDouble(wiersz.Cells[0].Value.ToString());
                    }
                }
                labelReg3V.Text = String.Format("3.3V: {0:0.00%} @ {1:00%}", currentMax, maxIndex);
                if (currentMax > 0.05) labelReg5V.ForeColor = Color.Red;
            }
            /////Spadek napięcia
            double drop12 = ((Convert.ToDouble(dataGridViewLoad.Rows[0].Cells[1].Value.ToString()) -
                Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[1].Value.ToString())) / 12);
            double drop5 = ((Convert.ToDouble(dataGridViewLoad.Rows[0].Cells[2].Value.ToString()) -
                Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[2].Value.ToString())) / 5);
            double drop3 = ((Convert.ToDouble(dataGridViewLoad.Rows[0].Cells[3].Value.ToString()) -
                Convert.ToDouble(dataGridViewLoad.Rows[dataGridViewLoad.Rows.Count - 1].Cells[3].Value.ToString())) / 3.3);

            labelDrop12.Text = String.Format("12V: {0:0.00%}", drop12);
            if (drop12 > 0.05) labelDrop12.ForeColor = Color.Red;

            labelDrop5.Text = String.Format("5V: {0:0.00%}", drop5);
            if (drop5 > 0.05) labelDrop5.ForeColor = Color.Red;

            labelDrop3.Text = String.Format("3.3V: {0:0.00%}", drop3);
            if (drop3 > 0.05) labelDrop3.ForeColor = Color.Red;

            double eff20 = 0.0, eff50 = 0.0, eff100 = 0.0;
            foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
            {
                if (Convert.ToDouble(wiersz.Cells[0].Value.ToString()) == 0.2) eff20 = Convert.ToDouble(wiersz.Cells[5].Value.ToString());
                else if (Convert.ToDouble(wiersz.Cells[0].Value.ToString()) == 0.5) eff50 = Convert.ToDouble(wiersz.Cells[5].Value.ToString());
                else if (Convert.ToDouble(wiersz.Cells[0].Value.ToString()) == 1) eff100 = Convert.ToDouble(wiersz.Cells[5].Value.ToString());
            }
            labelEffi20.Text = String.Format("20%: {0:00.00%}", eff20);
            labelEffi50.Text = String.Format("50%: {0:00.00%}", eff50);
            labelEffi100.Text = String.Format("100%: {0:00.00%}", eff100);

            if (eff20 > 0.8 && eff50 > 0.8 && eff100 > 0.8) labelEffiResult.Text = "= standard 80 Plus";
            if (eff20 > 0.82 && eff50 > 0.85 && eff100 > 0.82) labelEffiResult.Text = "= standard 80 Plus Bronze";
            if (eff20 > 0.85 && eff50 > 0.88 && eff100 > 0.85) labelEffiResult.Text = "= standard 80 Plus Silver";
            if (eff20 > 0.87 && eff50 > 0.9 && eff100 > 0.87) labelEffiResult.Text = "= standard 80 Plus Gold";
            if (eff20 > 0.9 && eff50 > 0.92 && eff100 > 0.89) labelEffiResult.Text = "= standard 80 Plus Platinum";
            if (eff20 > 0.92 && eff50 > 0.94 && eff100 > 0.9) labelEffiResult.Text = "= standard 80 Plus Titanium";
        }
        
        private void buttonWczytaj_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileLoaded = true;
                CSVdoDataTable(openFileDialog1.FileName);
                wypelnijLoadTable(dataTabelaLoad);
                wypelnijRippleTable(dataTabelaRipple);
                refreshAll();
            }
        }
        private void refreshAll()
        {
            formatujLoadTable();
            formatujRippleTable();
            wypelnijVregChart();
            wypelnijRippleChart();
            wypelnijNoiseChart();
            wypelnijPodsumowanie();
        }
        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            if (fileLoaded)
            {
                refreshAll();
                buttonRefresh.ForeColor = Color.Black;
            }
        }
        private void dataGridViewLoad_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            buttonRefresh.ForeColor = Color.Red;
        }

        private void radioButton12V_Click(object sender, EventArgs e) {wypelnijVregChart();}
        private void radioButton5V_Click(object sender, EventArgs e) {wypelnijVregChart();}
        private void radioButton3V_Click(object sender, EventArgs e) {wypelnijVregChart();}
        private void radioButton12VRipple_Click(object sender, EventArgs e) {wypelnijRippleChart();}
        private void radioButton5VRipple_Click(object sender, EventArgs e) {wypelnijRippleChart();}
        private void radioButton3VRipple_Click(object sender, EventArgs e) {wypelnijRippleChart();}

        private void buttonExportLoad_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string text = "<link rel=\"stylesheet\" type=\"text/css\" href=\"123.css\"><table class=\"psutable\">" +
                    "<colgroup><col class=\"load\"><col class=\"voltage\" span=\"3\"><col class=\"reszta\" span=\"4\"></colgroup><tr>";
                foreach (DataGridViewColumn header in dataGridViewLoad.Columns)
                {
                    text += "<th>" + header.Name + "</th>\n";
                }
                text += "</tr>";
                foreach (DataGridViewRow wiersz in dataGridViewLoad.Rows)
                {
                    text += "<tr>";
                    foreach (DataGridViewCell komorka in wiersz.Cells)
                    {
                        text += "<td>" + komorka.FormattedValue + "</td>\n";
                    }
                    text += "</tr>";
                }
                text += "</table>";

                File.WriteAllText(saveFileDialog1.FileName, text);
            }
        }
        private void buttonExportRipple_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string text = "<link rel=\"stylesheet\" type=\"text/css\" href=\"123.css\"><table class=\"psutable\">" +
                    "<colgroup><col class=\"load\"><col class=\"voltage\" span=\"3\"><tr>";
                foreach (DataGridViewColumn header in dataGridViewRipple.Columns)
                {
                    text += "<th>" + header.Name + "</th>\n";
                }
                text += "</tr>";
                foreach (DataGridViewRow wiersz in dataGridViewRipple.Rows)
                {
                    text += "<tr>";
                    foreach (DataGridViewCell komorka in wiersz.Cells)
                    {
                        text += "<td>" + komorka.FormattedValue + "</td>\n";
                    }
                    text += "</tr>";
                }
                text += "</table>";

                File.WriteAllText(saveFileDialog1.FileName, text);
            }
        }

        private void buttonImageVreg_Click(object sender, EventArgs e)
        {
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                chartVReg.SaveImage(saveFileDialog2.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
            }
        }
        private void buttonImageRipple_Click(object sender, EventArgs e)
        {
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                chartRipple.SaveImage(saveFileDialog2.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
            }
        }
        private void buttonImageNoise_Click(object sender, EventArgs e)
        {
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                chartNoise.SaveImage(saveFileDialog2.FileName, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
            }
        }
        
        private void buttonExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
