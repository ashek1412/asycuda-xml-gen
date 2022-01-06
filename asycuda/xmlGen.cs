using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace asycuda
{
    public partial class xmlGen : Form
    {
        public xmlGen()
        {
            InitializeComponent();
           // Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
           // Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
           // Microsoft.Office.Interop.Excel.Range range;
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Excel Files";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.DefaultExt = "xlsx";

         //   openFileDialog1.Filter = "Text files (*. OpenFileDialog openFileDialog1 = new OpenFileDialog()";

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;


            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "excel files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                textBox1.Text = openFileDialog1.FileName;
                buttonUpd.Enabled = true;

            }
        }

        private void buttonUpd_Click(object sender, EventArgs e)
        {


            try
            {
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + textBox1.Text + "';Extended Properties=Excel 12.0;");
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                dataGridView1.DataSource = DtSet.Tables[0];
                MyConnection.Close();




                if (dataGridView1.RowCount > 0 && validateExcel() == 1)
                    buttonXml.Enabled = true;
                else
                {
                    buttonXml.Enabled = false;
                    MessageBox.Show("Shipment data missing in excel file. Please Chcek !");
                }

            }
            catch (Exception ex)
            {

                string error = ex.Message;
            }



        }

        private void buttonXml_Click(object sender, EventArgs e)
        {
            if (textBoxFlt.Text.Trim().Length == 0)
            {
                MessageBox.Show("Enter Flight Number");
                return ;
            }
            if (textBoxMAWB.Text.Trim().Length == 0)
            {
                MessageBox.Show("Enter Master Airway Bill Number");
                return;
            }
            if (textBoxMAWB.Text.Trim().Length == 0)
            {
                MessageBox.Show("Enter Master Airway Bill Number");
                return;
            }

            string xml_name = "DEG101149007_" + DateTime.Now.ToString("dd") + DateTime.Now.ToString("MM") + DateTime.Now.ToString("yy") + DateTime.Now.ToString("HH")+ DateTime.Now.ToString("mm") + DateTime.Now.ToString("ss")+".xml";

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"C:\";
            saveFileDialog1.Title = "Save XML Files";
            saveFileDialog1.CheckFileExists = false;
            saveFileDialog1.CheckPathExists = true;
            saveFileDialog1.OverwritePrompt = true;
            saveFileDialog1.CreatePrompt = false;
            saveFileDialog1.DefaultExt = "xml";
            saveFileDialog1.Filter = "XML files (*.xml)|*.xml";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = xml_name;

           // XmlDocument doc = new XmlDocument();

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                XmlTextWriter writer = new XmlTextWriter(saveFileDialog1.FileName, System.Text.Encoding.UTF8);
            writer.WriteStartDocument(true);
            writer.Formatting = Formatting.Indented;
            writer.Indentation = 2;
            writer.WriteStartElement("Awbolds");
            writer.WriteStartElement("Master_bol");
            writer.WriteElementString("Customs_office_code", textBoxOC.Text.ToString());
            writer.WriteElementString("Voyage_number", textBoxFlt.Text.ToString());
            writer.WriteElementString("Date_of_departure", dateTimePicker1.Text.ToString());
            writer.WriteElementString("Reference_number", textBoxMAWB.Text.ToString());
            writer.WriteEndElement();

           

            //following will iterate all rows
            int Line_number = 1;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string Bol_reference = xml_entities(row.Cells[1].Value.ToString());
                string Goods_description = xml_entities(row.Cells[2].Value.ToString());
                string Port_of_origin_code = xml_entities(row.Cells[3].Value.ToString());
                string Carrier_code = xml_entities(row.Cells[4].Value.ToString());
                string Carrier_name = xml_entities(row.Cells[5].Value.ToString());
                string Carrier_address = xml_entities(row.Cells[6].Value.ToString());
                string Exporter_name = xml_entities(row.Cells[7].Value.ToString());
                string Exporter_address = xml_entities(row.Cells[8].Value.ToString());
                string Notify_code = xml_entities(row.Cells[9].Value.ToString());
                string Notify_name = xml_entities(row.Cells[10].Value.ToString());
                string Notify_address = xml_entities(row.Cells[11].Value.ToString());
                string Consignee_code = xml_entities(row.Cells[12].Value.ToString());
                string Consignee_name = xml_entities(row.Cells[13].Value.ToString());
                string Consignee_address = xml_entities(row.Cells[14].Value.ToString());
                int Number_of_packages =Int32.Parse(row.Cells[15].Value.ToString());
                string Package_type_code = xml_entities(row.Cells[16].Value.ToString());
                decimal Gross_mass = Math.Round(Decimal.Parse(row.Cells[17].Value.ToString()),2);
                string Shipping_marks = xml_entities(row.Cells[18].Value.ToString());

                writer.WriteStartElement("Bol_segment");
                    writer.WriteStartElement("Bol_id");
                        writer.WriteElementString("Bol_reference", Bol_reference);
                        writer.WriteElementString("Line_number", Line_number.ToString());
                        writer.WriteElementString("Bol_nature", "23");
                        writer.WriteElementString("Bol_type_code", "AWB");
                        writer.WriteElementString("Master_bol_ref_number", textBoxMAWB.Text.ToString());
                        writer.WriteElementString("DG_status", "");
                    writer.WriteEndElement();

                         writer.WriteElementString("Consolidated_Cargo", "0");

                    writer.WriteStartElement("Load_unload_place");
                        writer.WriteElementString("Port_of_origin_code", Port_of_origin_code);
                        writer.WriteElementString("Place_of_unloading_code", "BDDAC");
                    writer.WriteEndElement();

                    writer.WriteStartElement("Traders_segment");
                        writer.WriteStartElement("Carrier");
                              writer.WriteElementString("Carrier_code", Carrier_code);
                              writer.WriteElementString("Carrier_name", Carrier_name);
                              writer.WriteElementString("Carrier_address", Carrier_address);
                        writer.WriteEndElement();
                        writer.WriteStartElement("Shipping_Agent");
                                 writer.WriteElementString("Shipping_Agent_code", "");
                                 writer.WriteElementString("Shipping_Agent_name", "");
                        writer.WriteEndElement();
                        writer.WriteStartElement("Exporter");
                            writer.WriteElementString("Exporter_name", Exporter_name);
                            writer.WriteElementString("Exporter_address", Exporter_address);
                        writer.WriteEndElement();
                        writer.WriteStartElement("Notify");
                            writer.WriteElementString("Notify_code", "");
                            writer.WriteElementString("Notify_name", Notify_name);
                            writer.WriteElementString("Notify_address", Notify_address);
                       writer.WriteEndElement();
                       writer.WriteStartElement("Consignee");
                            writer.WriteElementString("Consignee_code","");
                            writer.WriteElementString("Consignee_name", Consignee_code);
                            writer.WriteElementString("Consignee_address", Consignee_address);
                       writer.WriteEndElement();
                  writer.WriteEndElement();

                  writer.WriteStartElement("Goods_segment");
                        writer.WriteElementString("Number_of_packages", "");
                        writer.WriteElementString("Package_type_code", Package_type_code);
                        writer.WriteElementString("Gross_mass", Gross_mass.ToString());
                        writer.WriteElementString("Shipping_marks", Shipping_marks);
                        writer.WriteElementString("Goods_description", Goods_description);
                        writer.WriteElementString("Volume_in_cubic_meters", "0.0");
                        writer.WriteElementString("Num_of_ctn_for_this_bol", "0");
                        writer.WriteElementString("Remarks", "");
                 writer.WriteEndElement();

                  writer.WriteStartElement("Value_segment");
                        writer.WriteStartElement("Freight_segment");
                            writer.WriteElementString("PC_indicator", "");
                            writer.WriteElementString("Freight_value", "0.0");
                            writer.WriteElementString("Freight_currency","ZZZ");
                        writer.WriteEndElement();
                        writer.WriteStartElement("Customs_segment");
                            writer.WriteElementString("Customs_value", "0.0");
                            writer.WriteElementString("Customs_currency", "");
                        writer.WriteEndElement();
                writer.WriteEndElement();



                writer.WriteEndElement();



                Line_number++;
            }

            writer.WriteEndElement();

            writer.WriteEndDocument();
            writer.Close();

             MessageBox.Show("XML File created ! ");



            }


        }


        private String xml_entities(String st)
        {
            String st1 = null;
            st1 = Regex.Replace(st, @"/[\x00-\x1F\x7f-\xFF]/", " ");
            st1 = Regex.Replace(st1,@"/[^A-Za-z0-9 _\-\+\&]/", " ");
            st1 = Regex.Replace(st1.Trim(),@"/[ ]{2,}|[\t]/", " " );

            return st1.Replace("<", " ").Replace(">", " ").Replace("\"", " ").Replace("'", " ").Replace("&", " and ").Replace("/", " ")
                        .Replace("\\", " ").Replace("%", " ").Replace("\\", " ").Replace("?", " ").Replace("`", " ").Replace("@", " ");

                    

        }



        private int validateExcel()
        {


            int isok = 1;
            int rok = 1;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                rok = 1;
                foreach (DataGridViewCell cl in row.Cells)
                {
                    if (string.IsNullOrEmpty(cl.Value.ToString()) && cl.ColumnIndex != 9 && cl.ColumnIndex != 12)
                    {
                        cl.Style.BackColor = System.Drawing.Color.Pink;
                        isok = 0;
                        rok = 0;

                    }
                }
                if (rok == 0)
                    row.Cells[0].Style.BackColor = System.Drawing.Color.Yellow;
                else
                    row.Cells[0].Style.BackColor = System.Drawing.Color.White;
            }

            return isok;



        }

    }


}
