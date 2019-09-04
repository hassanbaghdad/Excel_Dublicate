using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.Text.RegularExpressions;

namespace WindowsFormsApp11
{
    class Replacer
    {
        SQLiteConnection con = new SQLiteConnection(@"Data Source =replaces_db.db; Version=3;New=True;");
        string sql = "SELECT * FROM replaces_tbl";
        
        public string Filterd (string value)
        {

            SQLiteCommand cmd = new SQLiteCommand(sql, con);
            if (con.State != ConnectionState.Open)
            {
                con.Open();
            }
            SQLiteDataReader reader = cmd.ExecuteReader();
            
            while (reader.Read())
            {
                string Old_Value = reader[1].ToString();
                string New_Value = reader[2].ToString();
                value = value.Replace(Old_Value, New_Value);
            }
            for (int x = 0; x <= 60; x++)
            {

                if (value.Substring(0, 1) == " ")
                {

                    value = value.Remove(0, 1);

                }
            }



            return value.Replace("  ", string.Empty);
        }

    }
}
