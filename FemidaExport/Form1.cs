using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace FemidaExport
{
    public partial class Form1 : Form
    {
        string count = "0";
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add(@"C:\FEMIDA\Femida1");
            comboBox1.SelectedItem = comboBox1.Items[0];
            folderBrowserDialog1 = new FolderBrowserDialog();
            Console.WriteLine(comboBox1.SelectedItem);
            comboloader(comboBox1.SelectedItem.ToString());
            textBox2.ForeColor = Color.Gray;
            textBox2.Text = "Введите комментарий к проекту...";
            button2.Enabled = false;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            comboloader(comboBox1.SelectedItem.ToString());
        }
        private void textBox2_MouseClick(object sender, MouseEventArgs e)
        {
            if(textBox2.Text == "Введите комментарий к проекту...")
            {
                textBox2.Text = "";
                button2.Enabled = false;
                textBox2.ForeColor = Color.Black;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                comboBox1.Items.Add(folderBrowserDialog1.SelectedPath.ToString());
                comboBox1.SelectedItem = folderBrowserDialog1.SelectedPath;
            }
        }
        public void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            count = "0";
            string connectionString = null;
            connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Extended Properties=Paradox 3.x; Data Source=" + comboBox1.SelectedItem.ToString() + ";";
            Console.WriteLine(connectionString);
            OleDbConnection cnn;
            cnn = new OleDbConnection(connectionString);
            try
            {
                cnn.Open();
                Console.WriteLine("Select count(*) from Pr" + dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString());
                OleDbCommand command = new OleDbCommand(@"Select count(*) from Pr" + dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString().Replace(" ", ""), cnn);
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    count = reader[reader.FieldCount - 1].ToString();
                }
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            if (int.Parse(count) > 0 && textBox2.Text != "" && textBox2.Text != "Введите комментарий к проекту...")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }     
            label3.Text = "Количество элементов в выбранном проекте: " + count;
        }
        private string ConvertToUtf(string source)
        {
            source.Replace(";",",");
            Encoding srcEncodingFormat = Encoding.GetEncoding("windows-1252");
            byte[] originalByteString = srcEncodingFormat.GetBytes(source);
            return Encoding.Default.GetString(originalByteString);
        }
        private void comboloader(string path)
        {
            dataGridView1.Rows.Clear();
            if (path.Contains(@":\"))
            {
                string connectionString = null;
                connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Extended Properties=Paradox 3.x; Data Source=" + path + ";";
                Console.WriteLine(connectionString);
                OleDbConnection cnn;
                cnn = new OleDbConnection(connectionString);
                try
                {
                    cnn.Open();
                    OleDbCommand command = new OleDbCommand("Select Nomzak, Konstr, Podraz from Zakazi", cnn);
                    OleDbDataReader reader = command.ExecuteReader();
                    List<string[]> dataarr = new List<string[]>();
                    while (reader.Read())
                    {
                        dataarr.Add(new string[reader.FieldCount]);
                        dataarr[dataarr.Count - 1][0] = ConvertToUtf(reader[0].ToString());
                        dataarr[dataarr.Count - 1][1] = ConvertToUtf(reader[1].ToString());
                        dataarr[dataarr.Count - 1][2] = ConvertToUtf(reader[2].ToString());
                    }
                    reader.Close();
                    cnn.Close();
                    int t = 1;
                    foreach (string[] s in dataarr)
                    {
                        dataGridView1.Rows.Add(s);
                        dataGridView1.Rows[t-1].HeaderCell.Value = t.ToString();
                        t++;
                    }  
                    dataGridView1.RowHeadersWidth = t.ToString().Length * 25;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    MessageBox.Show("Проекты отсутствует");
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string connectionString = null;
            List<string> arrto = new List<string>();
            connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Extended Properties=Paradox 3.x; Data Source=" + comboBox1.SelectedItem.ToString() + ";";
            Console.WriteLine(connectionString);
            OleDbConnection cnn;
            cnn = new OleDbConnection(connectionString);
            try
            {
                cnn.Open();
                OleDbCommand command = new OleDbCommand("Select * from Pr"+ dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString(), cnn);
                OleDbDataReader reader = command.ExecuteReader();
                List<string[]> dataarr = new List<string[]>();
                while (reader.Read())
                {
                    dataarr.Add(new string[14]);
                    dataarr[dataarr.Count - 1][0] = Environment.MachineName.ToString();
                    dataarr[dataarr.Count - 1][1] = ConvertToUtf(reader[4].ToString());
                    dataarr[dataarr.Count - 1][2] = ConvertToUtf(reader[5].ToString());
                    dataarr[dataarr.Count - 1][3] = "Импортировано из фемиды";
                    dataarr[dataarr.Count - 1][4] = ConvertToUtf(reader[7].ToString());
                    dataarr[dataarr.Count - 1][5] = ConvertToUtf(reader[12].ToString());
                    dataarr[dataarr.Count - 1][6] = ConvertToUtf(reader[16].ToString()).Replace(",", ".");
                    dataarr[dataarr.Count - 1][7] = ConvertToUtf(reader[13].ToString()).Replace(",", ".");
                    dataarr[dataarr.Count - 1][8] = ConvertToUtf(reader[14].ToString()).Replace(",", ".");
                    dataarr[dataarr.Count - 1][9] = ConvertToUtf(reader[15].ToString()).Replace(",", ".");
                    dataarr[dataarr.Count - 1][10] = "0";
                    dataarr[dataarr.Count - 1][11] = textBox1.Text;
                    dataarr[dataarr.Count - 1][12] = Environment.UserName.ToString();
                    dataarr[dataarr.Count - 1][13] = "17";
                    string rowstoast = String.Join(";", dataarr[dataarr.Count - 1]);
                    arrto.Add(rowstoast);
                }
                Console.WriteLine(Path.GetTempPath());
                File.WriteAllLines(Path.GetTempPath() + "toAstr.txt", arrto.ToArray());
                reader.Close();
                cnn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                MessageBox.Show("Проекты отсутствует");
            }
            connectionString = @"Server=128.0.3.33; Port=3306; Database=femida; Uid=femida; Pwd=XXXXXXX; CharSet=utf8";
            MySqlConnection connection;
            connection = new MySqlConnection(connectionString);
            int startid = 1;
            try
            {
                connection.Open();
                Console.WriteLine("YEEES");
                MySqlCommand checkcomm = new MySqlCommand(String.Format("Select num from femida.projects where num like \"%{0}_FEMIDA%\"", dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString()), connection);
                MySqlDataReader reader = checkcomm.ExecuteReader();
                while (reader.Read())
                {
                    if (int.Parse(reader[0].ToString().Split('_')[2]) >= startid)
                    {
                        startid = int.Parse(reader[0].ToString().Split('_')[2]) + 1;
                    }
                    Console.WriteLine(reader[0].ToString() + "          tyt");
                }
                reader.Close();
                string[] arrtosql = new string[3];
                arrtosql[0] = String.Format("insert ignore into femida.projects(num, description, idx_projtype, glavkonstr, ggk, width, height, length) values('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}');", dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString() + "_FEMIDA_" + startid.ToString(), dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString() + " (Фемида) " + textBox2.Text, 3, dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString().Replace(" ", ""), dataGridView1[2, dataGridView1.CurrentRow.Index].Value.ToString().Replace(" ", ""), 100, 100, 100);
                arrtosql[1] = String.Format("insert ignore into femida.documents(idx_proj,num,name,date,idx_projtype,idx_doctype,user_name,user_desc,hostname)values(LAST_INSERT_ID(),'{0}(Фемида)','Импортировано из Фемиды, проект {1}',NOW(),'3','17','{2}','{3}','{4}'); SELECT LAST_INSERT_ID();", dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString(), dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString(), Environment.UserName.ToString(), textBox1.Text, Environment.MachineName.ToString());
                MySqlCommand command = new MySqlCommand(arrtosql[0] + arrtosql[1], connection);
                string idx = command.ExecuteScalar().ToString();
                foreach (string str in arrto)
                {
                    arrtosql[2] = arrtosql[2] + String.Format("insert ignore into femida.records(idx_doc,name,article_num_text,x,y,z,mass)values('{0}','{1}','{2}','{3}','{4}','{5}','{6}');", idx,str.Split(';')[4], str.Split(';')[5], str.Split(';')[7], str.Split(';')[8], str.Split(';')[9], str.Split(';')[6]);
                    //string[] strs = new string[str.Split(';').Length];
                }
                MySqlCommand comm = new MySqlCommand(arrtosql[2], connection);
                comm.ExecuteNonQuery();
                connection.Close();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
            }

        }
        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                textBox2.ForeColor = Color.Gray;
                textBox2.Text = "Введите комментарий к проекту...";
            }
            else
            {
                textBox2.ForeColor = Color.Black;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if ((textBox2.Text == "" || textBox2.Text == "Введите комментарий к проекту...") || int.Parse(count) <= 0)
            {
                button2.Enabled = false;
            }
            else
            {
                button2.Enabled = true;
            }
        }
    }
}
