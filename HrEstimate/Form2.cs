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
    public partial class Form2 : Form
    {
        XSC.LoaderFormInfo fFormInfo;
        XSC.ClientAccess.sqlClientAccess sca1;
        DataTable dt1;
        DataTable dt2;

        public Form2()
        {
            InitializeComponent();
        }

        private String v1;
        public String ppv1
        {
            get { return v1; }
            set { v1 = value; }
        }

        private String v2;
        public String ppv2
        {
            get { return v2; }
            set { v2 = value; }
        }

        private String v3;
        public String ppv3
        {
            get { return v3; }
            set { v3 = value; }
        }
        string v4;

        private void Form2_Load(object sender, EventArgs e)
        {
            fFormInfo = XSC.ClientLoader.FormInfo(this);

            sca1 = XSC.ClientAccess.UserAccess.sqlUserAccess(fFormInfo.LoginUserId);
            string sql = string.Format("select GROUPNAME from workitem_jack where wkclass='{0}'",v1);
            dt1 = sca1.GetDataTable(v2, sql);
            if (dt1.Rows.Count > 0)
            {
                v4 = dt1.Rows[0][0].ToString();
            }
            view1();
        }

        /// <summary>
        /// 顯示區域工作明細
        /// </summary>
        private void view1()
        {
            gridControl2.Visible = false;
            gridControl1.Visible = true;
            String strSQL = string.Format("select isnull(d.groupname,'其他') as 編制組別,c.GROUPNAME as 工作組別,wEmpID as 工號,wName as 姓名,case when c.GROUPNAME<>isnull(d.groupname,'') then '支援' else '' end as 備註,substring(convert(varchar,wSTime,120),12,5) as 開始時間 ", v1);
            strSQL = strSQL + string.Format(" from [192.168.100.208].[ERP].[dbo].[DC_EmpGroup] a right join Darren_RFWorkRecord_Agg b ", v1);
            strSQL = strSQL + string.Format("  on a.EmpID=b.wEmpID and convert(char(8),empdate,112)=convert(char(8),GETDATE(),112) join  workitem_jack c on replace(b.wZone,' ','')=replace(c.DCTYPE+c.GROUPNAME+c.ITEMNAME,' ','') ", v1);
            strSQL = strSQL + string.Format("  left join work_referce d on d.groupid=empgroupvalue where  convert(char(8),wSysTime,112)=convert(char(8),GETDATE(),112)  and wETime is null and wkclass='{0}' ", v1);
            dt1 = sca1.GetDataTable(v2, strSQL);
            gridControl1.DataSource = dt1;

            int[] width = new int[]{
                0,0,0,0,0,
                0,0,50,0,0,
                50,50
            };
            #region 計算支援人數
            int kk = 0;
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                if (dt1.Rows[i]["備註"].ToString() == "支援")
                {
                    kk = kk + 1;
                }
            }
            #endregion

            lbl_Content_Form2.Text = "支援"+v3+"工作" + kk.ToString() + "人 " + " 共 " + dt1.Rows.Count.ToString() + "人";

            lbl_WorkName_Form2.Text = v3;
        }

        /// <summary>
        /// 顯示編制人員工作明細
        /// </summary>
        private void view2()
        {
            gridControl2.Visible = true;
            gridControl1.Visible = false;
            gridControl1.DataSource = null;
            String strSQL1 = string.Format("select isnull(d.groupname,'其他') as 編制組別,c.GROUPNAME as 工作組別,wzone as 工作區域,wEmpID as 工號,wName as 姓名,case when c.GROUPNAME<>isnull(d.groupname,'') then '支援' else '' end as 備註,substring(convert(varchar,wSTime,120),12,5) as 開始時間 ", v1);
            strSQL1 = strSQL1 + string.Format(" from [192.168.100.208].[ERP].[dbo].[DC_EmpGroup] a right join Darren_RFWorkRecord_Agg b ", v1);
            strSQL1 = strSQL1 + string.Format("  on a.EmpID=b.wEmpID and convert(char(8),empdate,112)=convert(char(8),GETDATE(),112) join  workitem_jack c on replace(b.wZone,' ','')=replace(c.DCTYPE+c.GROUPNAME+c.ITEMNAME,' ','') ", v1);
            strSQL1 = strSQL1 + string.Format("  left join work_referce d on d.groupid=empgroupvalue where  convert(char(8),wSysTime,112)=convert(char(8),GETDATE(),112)  and  wetime  is null  and d.groupname='{0}'  ", v4);
            //    DataColumn dr_EstiHumanHr = new DataColumn("EstiHumanHr");
            //  dr_EstiHumanHr.DataType = System.Type.GetType("System.Int32");
            dt2 = sca1.GetDataTable(v2, strSQL1);
            gridControl2.DataSource = dt2;


            int[] width = new int[]{
                0,0,0,0,0,
                0,0,50,0,0,
                50,50
            };
            #region 計算支援人數
            int kk = 0;
             for (int i = 0; i < dt2.Rows.Count; i++)
            {
                if (dt2.Rows[i]["備註"].ToString() == "支援")
                {
                    kk = kk + 1;
                }
            }
            #endregion
             lbl_Content_Form2.Text = "支援他組" + kk.ToString() + "人 " + " 共 " + dt2.Rows.Count.ToString() + "人";

            lbl_WorkName_Form2.Text = v4;


              gridView2.ColumnPanelRowHeight = 40; // 標題列高
             
            for (int i = 0; i < dt2.Columns.Count; i++)
            {
                gridView2.Columns[i].BestFit();

                gridView2.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Top; // 至頂

                gridView2.Columns[i].AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap; // 縮小時往下直書

                try
                {
                    if (width[i] > 0)
                    {
                        gridView2.Columns[i].Width = width[i];
                    }
                }
                catch { }

            }  
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }

            if (e.KeyCode == Keys.F8)
            {
                this.Close();
            }

        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void btn_ShowDetail_Form2_Click(object sender, EventArgs e)
        {
            if (gridControl2.Visible == false)
            {
                view2();
                btn_ShowDetail_Form2.Text = "顯示區域工作明細";
            }
            else
            {
                view1();
                btn_ShowDetail_Form2.Text = "顯示編制人員工作明細";
            }
        }
    }
}
