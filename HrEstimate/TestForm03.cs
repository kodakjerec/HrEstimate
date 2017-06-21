using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HrEstimate
{
    public partial class TestForm03 : Form
    {
        //此表單的LoaderFormInfo
        XSC.LoaderFormInfo fFormInfo;

        //此LoginUserId所使用的sqlClientAccess
        XSC.ClientAccess.sqlClientAccess sca;


        public TestForm03()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            object[] obj_CheckAuthority = { "@a01", "NVarChar", "" };
            
            string strSelect = @"SELECT 
HAS_PERMS_BY_NAME(QUOTENAME(SCHEMA_NAME(schema_id)) + '.' + QUOTENAME(name), 'OBJECT','SELECT') AS select1, 
HAS_PERMS_BY_NAME(QUOTENAME(SCHEMA_NAME(schema_id)) + '.' + QUOTENAME(name), 'OBJECT','INSERT') AS insert1, 
HAS_PERMS_BY_NAME(QUOTENAME(SCHEMA_NAME(schema_id)) + '.' + QUOTENAME(name), 'OBJECT','DELETE') AS delete1, 
HAS_PERMS_BY_NAME(QUOTENAME(SCHEMA_NAME(schema_id)) + '.' + QUOTENAME(name), 'OBJECT','UPDATE') AS update1, 
* FROM sys.tables
order by name";
            DataTable dt_CheckTable = sca.GetDataTable(comboBox1.SelectedItem.ToString(), strSelect, obj_CheckAuthority, 0);
            dataGridView2.DataSource = dt_CheckTable;

            strSelect = 
@"select 
name
, has_perms_by_name(name, 'OBJECT', 'EXECUTE') as has_execute
, has_perms_by_name(name, 'OBJECT', 'VIEW DEFINITION') as has_view_definition 
from sys.procedures";
            DataTable dt_CheckSp = sca.GetDataTable(comboBox1.SelectedItem.ToString(), strSelect, obj_CheckAuthority, 0);
            dataGridView3.DataSource = dt_CheckSp;

            strSelect = 
@"select 
name
, has_perms_by_name(name, 'OBJECT', 'EXECUTE') as has_execute
, has_perms_by_name(name, 'OBJECT', 'VIEW DEFINITION') as has_view_definition
from sys.views";
            DataTable dt_CheckView = sca.GetDataTable(comboBox1.SelectedItem.ToString(), strSelect, obj_CheckAuthority, 0);
            dataGridView4.DataSource = dt_CheckView;

            strSelect = richTextBox1.Text;
            if (strSelect != "")
            {
                DataTable dt_CheckAuthority = sca.GetDataTable(comboBox1.SelectedItem.ToString(), strSelect, obj_CheckAuthority, 0);
                dataGridView1.DataSource = dt_CheckAuthority;
            }
        }

        private void TestForm03_Load(object sender, EventArgs e)
        {
            //取得此表單的LoaderFormInfo
            fFormInfo = XSC.ClientLoader.FormInfo(this);
            //透過LoginUserId取得sqlClientAccess
            sca = XSC.ClientAccess.UserAccess.sqlUserAccess(fFormInfo.LoginUserId);
            comboBox1.SelectedIndex = 0;
        }

        private void textBox1_Validated(object sender, EventArgs e)
        {
            button3.PerformClick();
        }
    }
}
