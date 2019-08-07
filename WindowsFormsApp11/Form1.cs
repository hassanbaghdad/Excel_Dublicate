using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Data.SQLite;

using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp11
{
    public partial class Form1 : Form
    {
        
        SQLiteConnection con = new SQLiteConnection(@"Data Source =ram_db.db; Version=3;New=True;");

        public string sheet_name;
        
        public Form1()
        {
            InitializeComponent();
             
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            
            }
           
        

        private void Button1_Click(object sender, EventArgs e)
        {

        }
       
        private void Button1_Click_1(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application ExcelObj = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook theWorkbook = null;
                string strPath = openFileDialog1.FileName;

                theWorkbook = ExcelObj.Workbooks.Open(strPath);

                Microsoft.Office.Interop.Excel.Sheets sheets = theWorkbook.Worksheets;
                for (int x = 1; x <= sheets.Count; x++)
                {
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(x);//Get the reference of second worksheet

                    sheet_name = worksheet.Name;//Get the name of worksheet.
                    listBox1.Items.Add(sheet_name);
                }
            }
           



        }

        private void Button3_Click(object sender, EventArgs e)
        {
            
        }

        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void ListBox1_SelectedValueChanged(object sender, EventArgs e)
        {
           

        }

        public void Button3_Click_1(object sender, EventArgs e)
        {
            
           
        }

        private void Button3_Click_2(object sender, EventArgs e)
        {
            try
            {
                con.Open();
                for (int x = 0; x <= listBox2.Items.Count - 1; x++)
                {
                    string fullName = listBox2.Items[x].ToString();
                    string[] names = new String[20];


                    names = fullName.Split(' ');

                    //1
                    if (names.Length == 1)
                    {
                        string fn = names[0];

                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");


                    }

                    // 2
                    if (names.Length == 2)
                    {

                        string fn = names[0];
                        string sn = names[1];

                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sn = sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");
                        string end_sn = Regex.Replace(re_sn, " {2,}", " ");

                        dataGridView2.Rows.Add(end_fn, end_sn);
                    }

                    // 3
                    if (names.Length == 3)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sn = sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_tn = tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");
                        string end_sn = Regex.Replace(re_sn, " {2,}", " ");
                        string end_tn = Regex.Replace(re_tn, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, null, null, null, null, null, null, null, listBox3.Items[x].ToString());



                    }
                    // 4
                    if (names.Length == 4)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sn = sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_tn = tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fon = fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");
                        string end_sn = Regex.Replace(re_sn, " {2,}", " ");
                        string end_tn = Regex.Replace(re_tn, " {2,}", " ");
                        string end_fon = Regex.Replace(re_fon, " {2,}", " ");


                        //string sql2 = "INSERT INTO names (fn  ) VALUES ('a')";
                        //SQLiteCommand cmd = new SQLiteCommand(sql2, con);

                        //cmd.ExecuteNonQuery();
                        ////, sn ,tn , fon ,fin ,sin ,sen ,ein ,nine, ten

                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, null, null, null, null, null, null, listBox3.Items[x].ToString());



                    }
                    // 5
                    if (names.Length == 5)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string fin = names[4];

                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sn = sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_tn = tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fon = fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fin = fin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");
                        string end_sn = Regex.Replace(re_sn, " {2,}", " ");
                        string end_tn = Regex.Replace(re_tn, " {2,}", " ");
                        string end_fon = Regex.Replace(re_fon, " {2,}", " ");
                        string end_fin = Regex.Replace(re_fin, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, end_fin, null, null, null, null, null, listBox3.Items[x].ToString());



                    }

                    // 6
                    if (names.Length == 6)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string fin = names[4];
                        string sin = names[5];

                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sn = sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_tn = tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fon = fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fin = fin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sin = sin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");
                        string end_sn = Regex.Replace(re_sn, " {2,}", " ");
                        string end_tn = Regex.Replace(re_tn, " {2,}", " ");
                        string end_fon = Regex.Replace(re_fon, " {2,}", " ");
                        string end_fin = Regex.Replace(re_fin, " {2,}", " ");
                        string end_sin = Regex.Replace(re_sin, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, end_fin, end_sin, null, null, null, null, listBox3.Items[x].ToString());




                    }

                    // 7
                    if (names.Length == 7)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string fin = names[4];
                        string sin = names[5];
                        string sen = names[6];

                        string re_fn = fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sn = sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_tn = tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fon = fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_fin = fin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sin = sin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        string re_sen = sen.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(re_fn, " {2,}", " ");
                        string end_sn = Regex.Replace(re_sn, " {2,}", " ");
                        string end_tn = Regex.Replace(re_tn, " {2,}", " ");
                        string end_fon = Regex.Replace(re_fon, " {2,}", " ");
                        string end_fin = Regex.Replace(re_fin, " {2,}", " ");
                        string end_sin = Regex.Replace(re_sin, " {2,}", " ");
                        string end_sen = Regex.Replace(re_sen, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, end_fin, end_sin, end_sen, null, null, null, listBox3.Items[x].ToString());




                    }
                    // 8
                    if (names.Length == 8)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string fin = names[4];
                        string sin = names[5];
                        string sen = names[6];
                        string ein = names[7];

                        fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        fin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sen.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        ein.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(fn, " {2,}", " ");
                        string end_sn = Regex.Replace(sn, " {2,}", " ");
                        string end_tn = Regex.Replace(tn, " {2,}", " ");
                        string end_fon = Regex.Replace(fon, " {2,}", " ");
                        string end_fin = Regex.Replace(fin, " {2,}", " ");
                        string end_sin = Regex.Replace(sin, " {2,}", " ");
                        string end_sen = Regex.Replace(sen, " {2,}", " ");
                        string end_ein = Regex.Replace(ein, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, end_fin, end_sin, end_sen, end_ein, null, null, listBox3.Items[x].ToString());




                    }
                    // 9
                    if (names.Length == 9)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string fin = names[4];
                        string sin = names[5];
                        string sen = names[6];
                        string ein = names[7];
                        string nin = names[8];

                        fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        fin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sen.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        ein.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        nin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(fn, " {2,}", " ");
                        string end_sn = Regex.Replace(sn, " {2,}", " ");
                        string end_tn = Regex.Replace(tn, " {2,}", " ");
                        string end_fon = Regex.Replace(fon, " {2,}", " ");
                        string end_fin = Regex.Replace(fin, " {2,}", " ");
                        string end_sin = Regex.Replace(sin, " {2,}", " ");
                        string end_sen = Regex.Replace(sen, " {2,}", " ");
                        string end_ein = Regex.Replace(ein, " {2,}", " ");
                        string end_nin = Regex.Replace(nin, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, end_fin, end_sin, end_sen, end_ein, end_nin, null, listBox3.Items[x].ToString());




                    }
                    // 10
                    if (names.Length == 10)
                    {

                        string fn = names[0];
                        string sn = names[1];
                        string tn = names[2];
                        string fon = names[3];
                        string fin = names[4];
                        string sin = names[5];
                        string sen = names[6];
                        string ein = names[7];
                        string nin = names[8];
                        string ten = names[9];

                        fn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        tn.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        fon.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        fin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        sen.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        ein.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        nin.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");
                        ten.Replace("  ", " ").Replace("ة", "ه").Replace("الله", " الله").Replace("عبد", "عبد ").Replace("ظ", "ض").Replace("أ", "ا").Replace("ؤ", "و").Replace("عبد", "عبد ").Replace("محمد", "محمد ");

                        string end_fn = Regex.Replace(fn, " {2,}", " ");
                        string end_sn = Regex.Replace(sn, " {2,}", " ");
                        string end_tn = Regex.Replace(tn, " {2,}", " ");
                        string end_fon = Regex.Replace(fon, " {2,}", " ");
                        string end_fin = Regex.Replace(fin, " {2,}", " ");
                        string end_sin = Regex.Replace(sin, " {2,}", " ");
                        string end_sen = Regex.Replace(sen, " {2,}", " ");
                        string end_ein = Regex.Replace(ein, " {2,}", " ");
                        string end_nin = Regex.Replace(nin, " {2,}", " ");
                        string end_ten = Regex.Replace(ten, " {2,}", " ");



                        dataGridView2.Rows.Add(end_fn, end_sn, end_tn, end_fon, end_fin, end_sin, end_sen, end_ein, end_nin, end_ten, listBox3.Items[x].ToString());




                    }


                }
            }
            catch
            {
                MessageBox.Show("يجب ان يحتوي الملف على الاسم الثلاثي على الاقل");
            }
           
        }
            private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (listBox1.Items.Count > 0)
                {
                    string selected_item = listBox1.SelectedItem.ToString();
                    OleDbConnection con = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0;Data Source = " + openFileDialog1.FileName + ";Extended Properties = 'Excel 12.0;HDR=YES';");
                    if (con.State != ConnectionState.Open)
                    {
                        con.Open();
                    }


                    string sql = "SELECT * FROM [" + selected_item + "$]";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(sql, con);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }
            catch
            {
                MessageBox.Show("رجاءا انقر مزدوجاً على اسم الشيت  ");
            }

           

        }

        public void Button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                int y = dataGridView1.Rows.Count - 1;
                string[] myarray = new string[y];

                for (int i = 0; i < y; i++)
                {
                    myarray[i] = dataGridView1.Rows[i].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString();

                    listBox2.Items.Add(myarray[i]);


                }
                label2.Text = listBox2.Items.Count.ToString();
            }
            catch
            {
                MessageBox.Show("تأكد من تحديد عمود الاسماء بشكل صحيح");
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView3.Rows.Clear();
                if (radioButton2.Checked)
                {
                    string fn = dataGridView2.Rows[0].Cells[0].Value.ToString();
                    string sn = dataGridView2.Rows[0].Cells[1].Value.ToString();
                    string tn = dataGridView2.Rows[0].Cells[2].Value.ToString();

                    int ts = 0;
                    for (int x = 0; x <= dataGridView2.Rows.Count - 2; x++)
                    {
                        int fn_rowindex = dataGridView2.Rows[x].Cells[0].RowIndex;
                        int sn_rowindex = dataGridView2.Rows[x].Cells[1].RowIndex;

                        int tn_rowindex = dataGridView2.Rows[x].Cells[2].RowIndex;

                        string curent_fn, curent_sn, curent_tn;
                        curent_fn = dataGridView2.Rows[x].Cells[0].Value.ToString();
                        curent_sn = dataGridView2.Rows[x].Cells[1].Value.ToString();
                        curent_tn = dataGridView2.Rows[x].Cells[2].Value.ToString();






                        for (int y = 0; y <= dataGridView2.Rows.Count - 2; y++)
                        {
                            if (dataGridView2.Rows[y].Cells[0].Value.ToString() == curent_fn & dataGridView2.Rows[y].Cells[1].Value.ToString() == curent_sn & dataGridView2.Rows[y].Cells[2].Value.ToString() == curent_tn
                           & fn_rowindex != dataGridView2.Rows[y].Cells[0].RowIndex & sn_rowindex != dataGridView2.Rows[y].Cells[1].RowIndex & tn_rowindex != dataGridView2.Rows[y].Cells[2].RowIndex

                          )
                            {

                                string a = dataGridView2.Rows[x].Cells[0].Value.ToString();
                                string b = dataGridView2.Rows[x].Cells[1].Value.ToString();
                                string c = dataGridView2.Rows[x].Cells[2].Value.ToString();
                                string d = dataGridView2.Rows[x].Cells[10].Value.ToString();

                                ts = ts + 1;
                                dataGridView3.Rows.Add(ts, a, b, c, d);

                            }
                        }

                        dataGridView3.Sort(dataGridView3.Columns[1], ListSortDirection.Ascending);
                    }

                    // listBox1.Items.Add();
                }



                //////////////////////////////////////////////////////
                ///


                if (radioButton1.Checked)
                {


                    string fn = dataGridView2.Rows[0].Cells[0].Value.ToString();
                    string sn = dataGridView2.Rows[0].Cells[1].Value.ToString();
                    //string tn = dataGridView2.Rows[0].Cells[2].Value.ToString();

                    int ts = 0;
                    for (int x = 0; x <= dataGridView2.Rows.Count - 2; x++)
                    {
                        int fn_rowindex = dataGridView2.Rows[x].Cells[0].RowIndex;
                        int sn_rowindex = dataGridView2.Rows[x].Cells[1].RowIndex;

                        //int tn_rowindex = dataGridView2.Rows[x].Cells[2].RowIndex;

                        string curent_fn, curent_sn, curent_tn;
                        curent_fn = dataGridView2.Rows[x].Cells[0].Value.ToString();
                        curent_sn = dataGridView2.Rows[x].Cells[1].Value.ToString();
                        // curent_tn = dataGridView2.Rows[x].Cells[2].Value.ToString();






                        for (int y = 0; y <= dataGridView2.Rows.Count - 2; y++)
                        {
                            if (dataGridView2.Rows[y].Cells[0].Value.ToString() == curent_fn & dataGridView2.Rows[y].Cells[1].Value.ToString() == curent_sn
                           & fn_rowindex != dataGridView2.Rows[y].Cells[0].RowIndex & sn_rowindex != dataGridView2.Rows[y].Cells[1].RowIndex

                          )
                            {

                                string a = dataGridView2.Rows[x].Cells[0].Value.ToString();
                                string b = dataGridView2.Rows[x].Cells[1].Value.ToString();
                                string c = dataGridView2.Rows[x].Cells[2].Value.ToString();
                                string d = dataGridView2.Rows[x].Cells[10].Value.ToString();

                                ts = ts + 1;
                                dataGridView3.Rows.Add(ts, a, b, c, d);

                            }
                        }

                        dataGridView3.Sort(dataGridView3.Columns[1], ListSortDirection.Ascending);
                    }

                    // listBox1.Items.Add();


                }
            }
            catch
            {
                MessageBox.Show("يوجد نقص بالاسماء الثلاثية");
            }
           
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void GroupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void Button5_Click_1(object sender, EventArgs e)
        {
            int rowindex = dataGridView1.CurrentCell.RowIndex;

            dataGridView1.Rows.RemoveAt(rowindex);
        }
        private void copyAlltoClipboard()
        {
            //to remove the first blank column from datagridview
            dataGridView3.RowHeadersVisible = false;
            dataGridView3.SelectAll();
            DataObject dataObj = dataGridView3.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void Button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;";

                saveFileDialog1.FileName = " مقاطــعة لـ " + Path.GetFileName(openFileDialog1.FileName);
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // creating Excel Application  
                    Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                    // creating new WorkBook within Excel application  
                    Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                    // creating new Excelsheet in workbook  
                    Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                    // see the excel sheet behind the program  
                    app.Visible = true;
                    // get the reference of first sheet. By default its name is Sheet1.  
                    // store its reference to worksheet  
                    worksheet = workbook.Sheets["ورقة1"];
                    worksheet = workbook.ActiveSheet;
                    // changing the name of active sheet  
                    worksheet.Name = "Exported from gridview";
                    // storing header part in Excel  
                    for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
                    }
                    // storing Each row and column value to excel sheet  
                    for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView3.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    // save the application  
                    workbook.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    // Exit from the application  
                    app.Quit();
                }

            }
            catch
            {
                MessageBox.Show("خطأ بالحفظ تأكد من مكان الحفظ واسم الملف");
            }
            
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            try
            {
                int y = dataGridView1.Rows.Count - 1;
                string[] myarray = new string[y];

                for (int i = 0; i < y; i++)
                {
                    myarray[i] = dataGridView1.Rows[i].Cells[dataGridView1.CurrentCell.ColumnIndex].Value.ToString();

                    listBox3.Items.Add(myarray[i]);





                }
                label4.Text = listBox3.Items.Count.ToString();
            }
            catch
            {
                MessageBox.Show("تأكد من تحديد عمود التشكيلات بشكل  صحيح ");
            }
           
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            dataGridView3.Rows.Clear();
        }

        private void ListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
    }

