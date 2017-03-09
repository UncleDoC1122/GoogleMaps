using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Maps.Geocoding;
using Google.Maps.StaticMaps;
using Google.Maps;
using System.Xml;
using Google_Maps_Address_Qualifier;

namespace SerialParser
{
    public partial class Form1 : Form
    {
        private string PathToInput = "D:\\input";
        private string PathToOutput = "D:\\output";
        private string PathToOutputUser = "";
        private string SaveType = "excelShort";
        private string InputType = "";


        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            PathToOutputUser = PathToOutput + ".xlsx";
            textBox2.Text = PathToOutputUser;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (PathToOutput.Equals("D:\\output") && !SaveType.Equals(""))
            {
                if (MessageBox.Show("Вы точно хотите сохранять в стандартную директорию?", "Предупреждение", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    textBox1.Text = Parsers.CheckAddress(textBox1.Text)[0].FormattedAddress;
                }
                else
                {
                    return;
                }
            }
            else if (SaveType.Equals(""))
            {
                MessageBox.Show("Выберите формат для сохранения!", "Ошибка!");
            }
            else
            {

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            switch (SaveType)
            {
                case "excelShort":
                    sfd.Filter = "Excel file(*.xlsx)|*.xlsx";
                    break;

                case "excelFull":
                    sfd.Filter = "Excel file(*.xlsx)|*.xlsx";
                    break;

                case "xmlShort":
                    sfd.Filter = "Xml file(*.xml)|*.xml";
                    break;

                case "xmlFull":
                    sfd.Filter = "Xml file(*.xml)|*.xml";
                    break;
                default:
                    MessageBox.Show("Выберите формат для сохранения!", "Ошибка!");
                    return;
                    break;
            }

            if (sfd.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            PathToOutput = sfd.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Убедитесь в том, что в исходном файле присутствуют соответствующие поля", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Csv file(*.csv)|*.csv|Excel file 1997-2003(*.xls)|*.xls|Excel file(*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            PathToInput = ofd.FileName;
            textBox2.Text = PathToInput;
            string[] tmp = ofd.SafeFileName.Split('.');
            string extension = tmp[tmp.Length - 1];

            switch (extension)
            {
                case "csv":
                    InputType = "csv";
                    break;
                case "xls":
                    InputType = "xls";
                    break;
                case "xlsx":
                    InputType = "xlsx";
                    break;
            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedItem.ToString())
            {
                case "Excel (только адреса и INTCODE)":
                    SaveType = "excelShort";
                    PathToOutputUser = PathToOutput + ".xlsx";
                    textBox2.Text = PathToOutputUser;
                    break;

                case "Excel (копировать исходный с вставкой в него адресов)":
                    SaveType = "excelFull";
                    PathToOutputUser = PathToOutput + ".xlsx";
                    textBox2.Text = PathToOutputUser;
                    break;

                case "Xml (только адреса и INTCODE)":
                    SaveType = "xmlShort";
                    PathToOutputUser = PathToOutput + ".xml";
                    textBox2.Text = PathToOutputUser;
                    break;

                case "Xml (копировать исходный со вставкой в него адресов)":
                    SaveType = "xmlFull";
                    PathToOutputUser = PathToOutput + ".xml";
                    textBox2.Text = PathToOutputUser;
                    break;
            }
        }
    }
}
