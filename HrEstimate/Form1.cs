using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Base;
using MailMachine;

namespace HrEstimate
{
    public partial class Form1 : Form
    {
        String Login_Server = "pxwms_n", SelStock = "DC01", DCName = "觀音倉";

        //此表單的LoaderFormInfo
        XSC.LoaderFormInfo fFormInfo;

        //此LoginUserId所使用的sqlClientAccess
        XSC.ClientAccess.sqlClientAccess sca;

        DataTable dt_Search, dt_SerachHistory, dtWk;
        DataTable AdminTable,   //XSC權限
                Table_SelStock, //廠內權限
                Table_Authority, //權限設定
                Table_RFbarcode, //RF工作項目
                Table_workreferce,  //組別維護
                Table_Authority_Del = new DataTable("Table_Authority_Del"),
                Table_RFbarcode_Del = new DataTable("Table_RFbarcode_Del");
        Random random = new Random();
        int Authority = 65535;
        int sk = 100;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //取得此表單的LoaderFormInfo
            fFormInfo = XSC.ClientLoader.FormInfo(this);
            //透過LoginUserId取得sqlClientAccess
            sca = XSC.ClientAccess.UserAccess.sqlUserAccess(fFormInfo.LoginUserId);

            #region 權限控制

            //先在XSC查詢工號和部門
            string AdminString = "select EMPID,ORG_ID,HECNAME from XSC_MENU_UserList where XSC_UserID=@userID";
            object[] objTable = { "@userID", "NChar", fFormInfo.UserId };
            AdminTable = sca.GetDataTable("EEPDC", AdminString, objTable, 0);
            label_UserID.Text = fFormInfo.UserId;
            label_EMPID.Text = AdminTable.Rows[0][0].ToString();
            label_Name.Text = AdminTable.Rows[0][2].ToString();

            #region 設定下拉式選單的艙別
            //取得EDI所設定的人效權限
            AdminString = "Select Lv,Site_no from PxWorkHrRole where Id=@userID";
            object[] objTable2 = { "@userID", "NChar", label_EMPID.Text };
            Table_SelStock = sca.GetDataTable("EDI_DC", AdminString, objTable2, 0);

            //倉別選擇
            #region 觀音乎?
            int SearchStock = MyDictionaryFind(Table_SelStock, "Site_no", "DC01");
            if (SearchStock >= 0)
            {
                comboBox_Stock.Items.Add(new Item("觀音", "pxwms_n"));
            }
            else
            {
                #region 檢查觀音WMS是否有權限
                string str_CheckAuthority = "Select top 1 0 from employee_data where S_empd_id=@userID";
                object[] obj_CheckAuthority = { "@userID", "NChar", label_EMPID.Text };
                DataTable dt_CheckAuthority = sca.GetDataTable("PXWMS_N", str_CheckAuthority, obj_CheckAuthority, 0);
                if (dt_CheckAuthority.Rows.Count > 0)
                {
                    string str_InsertAuthToEDI = "Insert into PxWorkHrRole values (@userID,@Name,'DC01','5','')";
                    object[] objTable3 = { "@userID", "NChar", label_EMPID.Text,
                                       "@Name","NChar",label_Name.Text};
                    sca.Update("EDI_DC", str_InsertAuthToEDI, objTable3, 0);

                    //新增入前台
                    object[] obj_NewSelStock = { 5, "DC01" };
                    Table_SelStock.Rows.Add(obj_NewSelStock);
                }
                #endregion
            }
            #endregion

            #region 岡山乎?
            SearchStock = -1;
            SearchStock = MyDictionaryFind(Table_SelStock, "Site_no", "DC02");
            if (SearchStock >= 0)
            {
                comboBox_Stock.Items.Add(new Item("岡山", "pxwms_s"));
            }
            else
            {
                #region 檢查岡山WMS是否有權限
                string str_CheckAuthority = "Select top 1 0 from employee_data where S_empd_id=@userID";
                object[] obj_CheckAuthority = { "@userID", "NChar", label_EMPID.Text };
                DataTable dt_CheckAuthority = sca.GetDataTable("PXWMS_S", str_CheckAuthority, obj_CheckAuthority, 0);
                if (dt_CheckAuthority.Rows.Count > 0)
                {
                    string str_InsertAuthToEDI = "Insert into PxWorkHrRole values (@userID,@Name,'DC02','5','')";
                    object[] objTable3 = { "@userID", "NChar", label_EMPID.Text,
                                       "@Name","NChar",label_Name.Text};
                    sca.Update("EDI_DC", str_InsertAuthToEDI, objTable3, 0);

                    //新增入前台
                    object[] obj_NewSelStock = { 5, "DC02" };
                    Table_SelStock.Rows.Add(obj_NewSelStock);
                }
                #endregion
            }
            #endregion

            #region 台中乎?
            SearchStock = -1;
            SearchStock = MyDictionaryFind(Table_SelStock, "Site_no", "DC03");
            if (SearchStock >= 0)
            {
                comboBox_Stock.Items.Add(new Item("梧棲", "pxwms_c_HumanHr"));
            }
            else
            {
                #region 檢查台中WMS是否有權限
                string str_CheckAuthority = "Select top 1 * from employee_data where S_empd_id=@userID";
                object[] obj_CheckAuthority = { "@userID", "NChar", label_EMPID.Text };
                DataTable dt_CheckAuthority = sca.GetDataTable("pxwms_c_HumanHr", str_CheckAuthority, obj_CheckAuthority, 0);
                if (dt_CheckAuthority.Rows.Count > 0)
                {
                    string str_InsertAuthToEDI = "Insert into PxWorkHrRole values (@userID,@Name,'DC03','5','')";
                    object[] objTable3 = { "@userID", "NChar", label_EMPID.Text,
                                       "@Name","NChar",label_Name.Text};
                    sca.Update("EDI_DC", str_InsertAuthToEDI, objTable3, 0);

                    //新增入前台
                    object[] obj_NewSelStock = { 5, "DC03" };
                    Table_SelStock.Rows.Add(obj_NewSelStock);
                }
                #endregion
            }
            #endregion
            #endregion

            Authority = Convert.ToInt32(Table_SelStock.Rows[0][0]);

            //判斷是否為資訊部人員
            if (AdminTable.Rows[0][1].ToString() == "970000")
                Authority = 0;

            if (Authority == 65535)
            {
                MessageBox.Show("您的工號 " + label_EMPID.Text + " 沒有使用權限\n請聯絡可設定權限人員", "權限不足");
                this.Close();
                return;
            }

            #region 關閉  人效/Hr 修改權限
            if (Authority >= 3)
            {
                colHumanHr.AppearanceCell.BackColor = Color.White;
                colHumanHr.OptionsColumn.AllowEdit = false;
                colHumanHr.OptionsColumn.AllowFocus = false;
            }
            #endregion

            #region 關閉 開始工作時間/休息時間/作業人數
            if (Authority >= 4)
            {
                //開始工作時間
                colStareTime.AppearanceCell.BackColor = Color.White;
                colStareTime.OptionsColumn.AllowEdit = false;
                colStareTime.OptionsColumn.AllowFocus = false;

                //休息時間起迄
                colRestSDate.AppearanceCell.BackColor = Color.White;
                colRestSDate.OptionsColumn.AllowEdit = false;
                colRestSDate.OptionsColumn.AllowFocus = false;
                colRestEDate.AppearanceCell.BackColor = Color.White;
                colRestEDate.OptionsColumn.AllowEdit = false;
                colRestEDate.OptionsColumn.AllowFocus = false;
                gridColumn1.OptionsColumn.AllowEdit = false;
                gridColumn2.OptionsColumn.AllowEdit = false;

                //作業人數
                colHuNum.AppearanceCell.BackColor = Color.White;
                colHuNum.OptionsColumn.AllowEdit = false;
                colHuNum.OptionsColumn.AllowFocus = false;
            }
            #endregion


            #endregion

            #region 下拉式選單

            #region 頁面_查詢_收費類別
            DataTable dt0 = sca.GetDataTable(Login_Server, "select WClass,WClassNm from Px_HrSet order by Sn");
            DataRow BeginRow = dt0.NewRow();
            BeginRow["WClass"] = "";
            BeginRow["WClassNm"] = "全部";
            dt0.Rows.Add(BeginRow);

            comboBox_Search.DisplayMember = "WClassNm";
            comboBox_Search.ValueMember = "WClass";
            comboBox_Search.DataSource = dt0;
            comboBox_Search.SelectedIndex = dt0.Rows.Count - 1;
            #endregion

            #region 頁面_歷史查詢_收費類別
            DataTable dt1 = dt0.Copy();
            comboBox_SearchHistory.DisplayMember = "WClassNm";
            comboBox_SearchHistory.ValueMember = "WClass";
            comboBox_SearchHistory.DataSource = dt1;
            comboBox_SearchHistory.SelectedIndex = dt1.Rows.Count - 1;
            #endregion

            #region 頁面_RF權限_組別
            button_workreferce_Query();
            #endregion

            #endregion

            gridColumn40.DisplayFormat.Format = new XXXXFormatter();
            tabControl1.SelectedIndex = 1;
            comboBox_Stock.SelectedIndex = 0;
            comboBox_Stock_SelectedIndexChanged(sender, e);
            tabControl1_SelectedIndexChanged(sender, e);

            #region Settings的刪除Table
            Settings_Detail_Del();
            #endregion

            #region 權限設定的刪除Table
            Table_Authority_Del.Columns.Add("Id", typeof(String));
            Table_Authority_Del.Columns.Add("Site_no", typeof(String));
            #endregion

            #region RF條碼設定的刪除Table
            Table_RFbarcode_Del.Columns.Add("barcode", typeof(String));
            #endregion
        }

        #region 頁面_設定
        /// <summary>
        /// 修改Cell時馬上紀錄時間
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gridView_Settings_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Value != null)
            {
                gridView_Settings.SetFocusedRowCellValue("Up_date", DateTime.Now);
                gridView_Settings.SetFocusedRowCellValue("UpUser", fFormInfo.LoginUserId);
            }
        }

        /// <summary>
        /// 按下查詢按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Settings_Query()
        {
            string cmdstring = "select Sn,WClass,WClassNm,HumanHr,StareTime,RestTime,HuNum,Up_date,UpUser from Px_HrSet";
            ds_Settings.Clear();
            dt_Settings.Merge(sca.GetDataTable(Login_Server, cmdstring));

            cmdstring = "select Sn,WClass,RestS,RestE,RestMin,RestSDate=convert(datetime,convert(date,GETDATE()))+RestS,RestEDate=convert(datetime,convert(date,GETDATE()))+RestE "
                    + "from Px_HrRestSet";
            dt_RestTime.Clear();
            dt_RestTime.Merge(sca.GetDataTable(Login_Server, cmdstring));
            gridView_Settings.CollapseAllDetails();
        }

        #region 休息時間起迄日期檢查
        /// <summary>
        /// 檢查休息時間區間
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Settings_Detail_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {
            DataRowView drv = e.Row as DataRowView;
            drv["RestS"] = new TimeSpan(DateTime.Parse(drv["RestSDate"].ToString()).Hour,
                                        DateTime.Parse(drv["RestSDate"].ToString()).Minute,
                                        DateTime.Parse(drv["RestSDate"].ToString()).Second);
            drv["RestE"] = new TimeSpan(DateTime.Parse(drv["RestEDate"].ToString()).Hour,
                                        DateTime.Parse(drv["RestEDate"].ToString()).Minute,
                                        DateTime.Parse(drv["RestEDate"].ToString()).Second);
            TimeSpan restS = TimeSpan.Parse(drv["RestS"].ToString()),
                     restE = TimeSpan.Parse(drv["RestE"].ToString());
            TimeSpan passtime = restE - restS;

            if (passtime.TotalMinutes < 0)
            {
                e.Valid = false;
                e.ErrorText = "起>迄";
            }

            drv["RestMin"] = passtime.TotalMinutes;
            Calculate_RestTime(drv["WClass"].ToString());
        }

        void Calculate_RestTime(string WClassName)
        {
            double TotalRestTime = 0;
            foreach (DataRow dr in dt_RestTime.Rows)
            {
                if (dr.RowState != DataRowState.Deleted)
                    if (dr["WClass"].ToString() == WClassName)
                        TotalRestTime += Convert.ToDouble(dr["RestMin"]);
            }
            int Newrowhandle = gridView_Settings.LocateByValue(0, colWClass, WClassName);

            //Get Row index
            int RowIndex = gridView_Settings.GetDataSourceRowIndex(Newrowhandle);
            if (dt_Settings.Rows[RowIndex].RowState == DataRowState.Unchanged)
                dt_Settings.Rows[RowIndex].SetModified();

            gridView_Settings.SetRowCellValue(Newrowhandle, colRestTime, TotalRestTime);
        }

        /// <summary>
        /// 輸入Setting_Detail錯誤
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Settings_Detail_InvalidRowException(object sender, DevExpress.XtraGrid.Views.Base.InvalidRowExceptionEventArgs e)
        {
            e.ExceptionMode = ExceptionMode.Ignore;
            MessageBox.Show("起迄日區間設定錯誤", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);

            Settings_Detail.HideEditor();
        }
        #endregion

        #region 休息時間新增
        /// <summary>
        /// 新增休息時間明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemButtonEdit_Settings_Add_Click(object sender, EventArgs e)
        {
            DataRowView drv = gridView_Settings.GetFocusedRow() as DataRowView;
            string WClass = drv["WClass"].ToString();

            DataRow Newdr = dt_RestTime.NewRow();

            Newdr["Sn"] = random.Next();
            Newdr["WClass"] = WClass;
            Newdr["RestS"] = "09:00:00";
            Newdr["RestE"] = "09:00:00";
            Newdr["RestMin"] = 0;
            Newdr["RestSDate"] = DateTime.Now.ToShortDateString() + " 09:00:00";
            Newdr["RestEDate"] = DateTime.Now.ToShortDateString() + " 09:00:00";

            dt_RestTime.Rows.Add(Newdr);
            Calculate_RestTime(WClass);
        }

        #region 透過Detail的欄位標題新增
        ///// <summary>
        ///// 透過Detail的欄位標題新增
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void Settings_Detail_MouseDown(object sender, MouseEventArgs e)
        //{
        //    //取得滑鼠座標
        //    Point p = gridControl_Settings.PointToClient(MousePosition);
        //    BaseView view = gridControl_Settings.FocusedView;
        //    DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo hi =
        //        view.CalcHitInfo(p) as DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo;
        //    if (hi.InRow && hi.Column == null)
        //    {
        //        DataRow dr = (view as DevExpress.XtraGrid.Views.Grid.GridView).GetDataRow(hi.RowHandle);
        //        string WClassName = dr["WClass"].ToString();
        //        int WClass_RowCount = 0;
        //        foreach (DataRow dr1 in dt_RestTime.Rows)
        //        {
        //            if (dr1["WClass"].ToString() == WClassName)
        //            {
        //                WClass_RowCount++;
        //            }
        //        }
        //        if (WClass_RowCount == 0)
        //        {
        //            EventArgs e1 = new EventArgs();
        //            repositoryItemButtonEdit1_Click(sender, e1);
        //        }
        //    }
        //}
        #endregion

        /// <summary>
        /// Master無資料自動新增
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gridView_Settings_MasterRowExpanded(object sender, CustomMasterRowEventArgs e)
        {
            BaseView bv = gridView_Settings.GetDetailView(gridView_Settings.FocusedRowHandle, 0);
            if (bv != null)
            {
                if (bv.DataRowCount == 0)
                {
                    EventArgs e1 = new EventArgs();
                    repositoryItemButtonEdit_Settings_Add_Click(sender, e1);
                }
            }
        }
        #endregion

        #region 休息時間刪除
        DataTable dt_Settings_Del = new DataTable();
        /// <summary>
        /// 建立休息時間刪除Table
        /// </summary>
        private void Settings_Detail_Del()
        {
            DataColumn dc = new DataColumn("DelSn", System.Type.GetType("System.Int32"));
            dt_Settings_Del.Columns.Add(dc);
            DataColumn dc1 = new DataColumn("WClass", System.Type.GetType("System.String"));
            dt_Settings_Del.Columns.Add(dc1);
        }
        /// <summary>
        /// 刪除休息時間明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemButtonEdit_Settings_Del_Click(object sender, EventArgs e)
        {
            string WClassName = "";

            //將刪除資料加進Del Table
            DataRow dr = dt_Settings_Del.NewRow();

            //取得滑鼠座標
            Point p = gridControl_Settings.PointToClient(MousePosition);
            BaseView view = gridControl_Settings.FocusedView;
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo hi =
                view.CalcHitInfo(p) as DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo;
            if (hi.InRow)
            {
                int rowHandle = hi.RowHandle;
                // do something;
                DataRow row = (view as DevExpress.XtraGrid.Views.Grid.GridView).GetDataRow(rowHandle);
                dr[0] = Convert.ToInt32(row["Sn"]);
                dr[1] = row["WClass"].ToString();
                WClassName = row["WClass"].ToString();
            }

            dt_Settings_Del.Rows.Add(dr);

            controlNavigator1.Buttons.DoClick(controlNavigator1.Buttons.EndEdit);
            controlNavigator1.Buttons.DoClick(controlNavigator1.Buttons.Remove);

            //重新計算休息時間
            Calculate_RestTime(WClassName);
        }
        #endregion

        /// <summary>
        /// 按下儲存按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Settings_Save_Click(object sender, EventArgs e)
        {
            #region 是否有修改?
            int Modify_Count = 0;
            DataTable Modify_Table = ds_Settings.Tables[0],
                      Modify_Table2 = ds_Settings.Tables[1];
            if (Modify_Table == null || Modify_Table2 == null)
            {
                MessageBox.Show("請先做查詢:)");
                return;
            }

            Modify_Table.TableName = "Px_HrSet";
            Modify_Table2.TableName = "Px_HrRestSet";

            for (int i = Modify_Table.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = Modify_Table.Rows[i];
                if (dr.RowState != DataRowState.Unchanged)
                    Modify_Count++;
            }
            for (int i = Modify_Table2.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = Modify_Table2.Rows[i];
                if (dr.RowState != DataRowState.Unchanged)
                    Modify_Count++;
            }
            #endregion

            #region 開始回寫資料
            if (Modify_Count > 0)
            {
                DialogResult dr1 = MessageBox.Show("確定修改資料??", "修改確認", MessageBoxButtons.YesNo);
                if (dr1 == System.Windows.Forms.DialogResult.No)
                    return;
                //儲存資料
                int SuccessUpdateCount_Master = 0;
                string Update_string = "Update Px_HrSet set [HuNum]=@HuNum,[HumanHr]=@HumanHr,[StareTime]=@StareTime,[RestTime]=@RestTime,"
                                    + "Up_date=getdate(),UpUser=@LoginUserId where [WClass]=@WClass;Select SuccessCount=@@rowcount";
                string Detail_Update = "Update Px_HrRestSet set RestS=@RestS,RestE=@RestE,RestMin=@RestMin where Sn=@Sn;Select SuccessCount=@@rowcount";
                string Detail_Insert = "Insert Into Px_HrRestSet values (@WClass,@RestS,@RestE,@RestMin);Select SuccessCount=@@rowcount";
                string Detail_Delete = "Delete Px_HrRestSet Where Sn=@Sn;Select SuccessCount=@@rowcount";
                string Result_string = "";

                #region Master
                for (int i = Modify_Table.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = Modify_Table.Rows[i];
                    if (dr.RowState == DataRowState.Modified)
                    {
                        object[] obj_Param = { "@HuNum", "Decimal", dr["HuNum"],
                                               "@HumanHr", "Decimal", dr["HumanHr"],
                                               "@StareTime", "Time", dr["StareTime"],
                                               "@RestTime", "Decimal", dr["RestTime"],
                                               "@WClass", "NVarChar", dr["WClass"],
                                               "@LoginUserId","NVarChar",fFormInfo.LoginUserId};
                        DataTable test1 = sca.GetDataTable(Login_Server, Update_string, obj_Param, 0);
                        SuccessUpdateCount_Master += Convert.ToInt32(test1.Rows[0]["SuccessCount"]);
                        if (Convert.ToInt32(test1.Rows[0]["SuccessCount"]) > 0)
                            Result_string += "【" + dr["WClassNm"].ToString() + "】"
                                + " 成功更新"
                                + " 人效/Hr：" + dr["HumanHr"].ToString()
                                + " 開始時間：" + dr["StareTime"].ToString()
                                + " 休息時間：" + dr["RestTime"].ToString()
                                + " 作業人數：" + dr["HuNum"].ToString() + "\r\n";
                    }
                }
                #endregion

                #region Detail
                for (int i = Modify_Table2.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = Modify_Table2.Rows[i];
                    if (dr.RowState == DataRowState.Deleted)
                        continue;

                    object[] obj_Param = { "@RestS", "Time", dr["RestS"],
                                           "@RestE", "Time", dr["RestE"],
                                           "@RestMin", "Int", dr["RestMin"],
                                           "@Sn", "Int", dr["Sn"],
                                           "@WClass","NVarChar", dr["WClass"]};

                    #region Detail-Update
                    if (dr.RowState == DataRowState.Modified)
                    {
                        DataTable test1 = sca.GetDataTable(Login_Server, Detail_Update, obj_Param, 0);
                        SuccessUpdateCount_Master += Convert.ToInt32(test1.Rows[0]["SuccessCount"]);
                        if (Convert.ToInt32(test1.Rows[0]["SuccessCount"]) > 0)
                            Result_string += "【" + dr["WClass"].ToString() + "】"
                                + " 成功更新休息時間：" + dr["RestS"].ToString()
                                + " ~ " + dr["RestE"].ToString()
                                + " 經過時間：" + dr["RestMin"].ToString() + "\r\n";
                    }
                    #endregion

                    #region Detail-Add
                    if (dr.RowState == DataRowState.Added)
                    {
                        DataTable test1 = sca.GetDataTable(Login_Server, Detail_Insert, obj_Param, 0);
                        SuccessUpdateCount_Master += Convert.ToInt32(test1.Rows[0]["SuccessCount"]);
                        if (Convert.ToInt32(test1.Rows[0]["SuccessCount"]) > 0)
                            Result_string += "【" + dr["WClass"].ToString() + "】"
                                + " 成功新增休息時間：" + dr["RestS"].ToString()
                                + " ~ " + dr["RestE"].ToString()
                                + " 經過時間：" + dr["RestMin"].ToString() + "\r\n";
                    }
                    #endregion
                }

                #region Detail-Del

                for (int i = dt_Settings_Del.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = dt_Settings_Del.Rows[i];
                    object[] obj_Param = { "@Sn", "Int", dr["DelSn"] };

                    DataTable test1 = sca.GetDataTable(Login_Server, Detail_Delete, obj_Param, 0);
                    SuccessUpdateCount_Master += Convert.ToInt32(test1.Rows[0]["SuccessCount"]);
                    if (Convert.ToInt32(test1.Rows[0]["SuccessCount"]) > 0)
                        Result_string += "【" + dr["WClass"].ToString() + "】"
                            + " 成功刪除休息時間\r\n";
                }
                #endregion

                #endregion

                if (SuccessUpdateCount_Master > 0)
                {
                    Modify_Table.AcceptChanges();
                    Modify_Table2.AcceptChanges();
                    MessageBox.Show(Result_string, "更新成功");
                }
            }
            else
            {
                MessageBox.Show("無資料被修改");
            }
            #endregion

            #region 重新查詢
            button_Settings_Query();
            #endregion
        }
        #endregion

        #region 頁面_查詢

        #region gridView_Search重算功能

        /// <summary>
        /// 按下Enter後啟動認證程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gridView_Search_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            #region 預算功能相關變數
            TimeSpan span;      //耗費時間_時分秒
            double NewCapacity = 0, //新產能
                spanHours = 0,  //耗費時(hr)
                spanMin = 0,    //耗費分(min)
                spanRest = 0;   //耗費休息時間(min)

            int NewEstiQty = 0; //新預估作業量

            string StartDateTime;   //目前時間
            DateTime NewEstiTime = DateTime.Now;    //新預估完成時間

            string NewStatus;      //新預估狀態
            double NewStatusCalA = 0,//新預估狀態_前導計算值
                NewStatusCalB = 0;//新預估狀態_對照值
            #endregion

            DataRow dr = gridView_Search.GetFocusedDataRow();

            //新產能
            NewCapacity = (double)dr["HumanHr"] * (int)dr["HuNum"];
            if (NewCapacity == 0)
                return;

            StartDateTime = DateTime.Parse(dr["WDate"].ToString()).ToShortDateString()
                + " " + dr["StareTime"].ToString();
            span = DateTime.Now.Subtract(DateTime.Parse(StartDateTime));
            spanHours = span.Hours;
            spanMin = Convert.ToDouble(span.Minutes);
            spanRest = Convert.ToDouble(dr["RestTime"]);

            //新預估作業量
            //早上:產能*經過時間
            if (DateTime.Now.Hour < 12)
            {
                NewEstiQty = Convert.ToInt32(Math.Ceiling(NewCapacity * (spanHours + spanMin / 60)));
            }
            //下午:產能*(經過時間-休息時間)
            else
            {
                NewEstiQty = Convert.ToInt32(Math.Ceiling(NewCapacity * (spanHours + spanMin / 60 - spanRest / 60)));
            }

            //新預估完成時間
            //早上:(剩餘作業量 / 產能)*60+(現在時間+休息時間)
            if (DateTime.Now.Hour < 12)
            {
                NewEstiTime = DateTime.Now.AddMinutes(spanRest).AddMinutes((Convert.ToDouble(dr["剩餘作業量"]) / NewCapacity) * 60);
            }
            //下午:(剩餘作業量 / 產能)*60+(現在時間)
            else
            {
                NewEstiTime = DateTime.Now.AddMinutes((Convert.ToDouble(dr["剩餘作業量"]) / NewCapacity) * 60);
            }
            if (NewEstiTime < DateTime.Now)
            {
                NewEstiTime = DateTime.Now;
            }

            //新預估狀態
            NewStatusCalB = Convert.ToDouble(dr["實際作業量"]) / Convert.ToDouble(dr["CountAll"]);

            if (Convert.ToInt32(dr["EstiQty"]) == 0 ||
                (Convert.ToDouble(dr["實際作業量"]) / Convert.ToDouble(dr["CountAll"]) > 1))
                NewStatusCalA = 0;
            else
                NewStatusCalA = NewEstiQty / Convert.ToDouble(dr["CountAll"]);
            if (NewStatusCalA > NewStatusCalB)
                NewStatus = "延遲";
            else
                NewStatus = "準時";


            if (NewStatusCalA > NewStatusCalB)
                NewStatus = "延遲";
            else
                NewStatus = "準時";


            if (DateTime.Compare(NewEstiTime, DateTime.Parse(dr["原預計完成時間"].ToString())) == 0)
            {
                NewStatus = "準時";
            }
            else if ((DateTime.Compare(NewEstiTime, DateTime.Parse(dr["原預計完成時間"].ToString())) < 0))
            {
                NewStatus = "提早";
            }
            else if ((DateTime.Compare(NewEstiTime, DateTime.Parse(dr["原預計完成時間"].ToString())) > 0))
            {
                NewStatus = "延遲";
            }



            dr["Capacity"] = NewCapacity;
            dr["預計產值"] = NewEstiQty;
            dr["預計完成時間"] = NewEstiTime;
            dr["狀態"] = NewStatus;
        }
        #endregion

        /// <summary>
        /// 作業類別切換
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox_Search_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Search_WClass = comboBox_Search.SelectedValue == null ? "" : comboBox_Search.SelectedValue.ToString();
            string FilterString = "";

            if (Search_WClass == "")
                FilterString = "";
            else
                FilterString = "[WClass]='" + Search_WClass + "'";

            gridView_Search.ActiveFilterString = FilterString;
        }

        /// <summary>
        /// 查詢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Search_Click(object sender, EventArgs e)
        {
            button_Search.Enabled = false;
            //步驟一：先執行sp
            object[] obj_Table = { "@test", "NVarChar", "test" };
            this.Cursor = Cursors.WaitCursor;
            dt_Search = sca.GetDataTable(Login_Server, "exec sp_PxImdWork");
            DataColumn dr_EstiHumanHr = new DataColumn("EstiHumanHr");
            dr_EstiHumanHr.DataType = System.Type.GetType("System.Int32");
            dt_Search.Columns.Add(dr_EstiHumanHr);
            gridControl_Search.DataSource = dt_Search;

            #region 找到[XD_收貨]的實際作業量,取代[Sorter裝籠]的實際作業量
            //int XD_ReciBox_seq = MyDictionaryFind(dt_Search, "WClass", "XD_ReciBox");
            //if (XD_ReciBox_seq >= 0)
            //{
            //    object XD_ReciBox_RealQty = dt_Search.Rows[XD_ReciBox_seq]["實際作業量"];
            //    int Sorter_seq = MyDictionaryFind(dt_Search, "WClass", "Sorter");
            //    dt_Search.Rows[Sorter_seq]["實際作業量"] = XD_ReciBox_RealQty;
            //}
            #endregion

            gridControl_Search.Focus();
            this.Cursor = Cursors.Default;
            button_Search.Enabled = true;
        }

        /// <summary>
        /// 確認
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Confirm_Click(object sender, EventArgs e)
        {
            int Modify_Count = 0;
            //是否有修改?
            DataTable Modify_Table = (DataTable)gridControl_Search.DataSource;
            if (Modify_Table == null)
            {
                MessageBox.Show("請先做查詢:)");
                return;
            }
            for (int i = Modify_Table.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = Modify_Table.Rows[i];
                if (dr.RowState == DataRowState.Modified)
                    Modify_Count++;
            }

            if (Modify_Count > 0)
            {
                DialogResult dr1 = MessageBox.Show("確定修改資料??", "修改確認", MessageBoxButtons.YesNo);
                if (dr1 == System.Windows.Forms.DialogResult.No)
                    return;
                //儲存資料
                int SuccessUpdateCount = 0;
                string Update_string = "Update Px_HrSet set [HuNum]=@HuNum,Up_date=getdate(),UpUser=@LoginUserId where [WClass]=@WClass;Select SuccessCount=@@rowcount";
                string Result_string = "";
                for (int i = Modify_Table.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = Modify_Table.Rows[i];
                    if (dr.RowState == DataRowState.Modified)
                    {
                        object[] obj_Param = { "@HuNum", "Decimal", dr["HuNum"],
                                               "@WClass", "NVarChar", dr["WClass"],
                                               "@LoginUserId","NVarChar",fFormInfo.LoginUserId};
                        DataTable test1 = sca.GetDataTable(Login_Server, Update_string, obj_Param, 0);
                        SuccessUpdateCount += Convert.ToInt32(test1.Rows[0]["SuccessCount"]);
                        if (Convert.ToInt32(test1.Rows[0]["SuccessCount"]) > 0)
                            Result_string += dr["WClassNm"].ToString() + " 成功更新作業人數為 " + dr["HuNum"].ToString() + "\n";
                    }
                }
                if (SuccessUpdateCount > 0)
                {
                    Modify_Table.AcceptChanges();
                    MessageBox.Show(Result_string, "更新成功");
                }
            }
            else
            {
                MessageBox.Show("無資料被修改");
            }
        }

        /// <summary>
        /// 點選作業量跳出該工作目前人員數量
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gridView_Search_RowClick(object sender, RowClickEventArgs e)
        {
            // MessageBox.Show(gridView_Search.GetRowCellDisplayText(e.RowHandle, gridView_Search.Columns[1]));

            Form2 pop = new Form2();




            pop.ppv1 = gridView_Search.GetRowCellDisplayText(e.RowHandle, gridView_Search.Columns[1]);

            pop.ppv2 = Login_Server;
            pop.ppv3 = gridView_Search.GetRowCellDisplayText(e.RowHandle, gridView_Search.Columns[2]);

            pop.ShowDialog();


        }

        #endregion

        #region 頁面_歷史查詢
        /// <summary>
        /// 查詢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_SearchHistory_Click(object sender, EventArgs e)
        {
            string BDate = dateTimePicker_BDate.Value.ToString("yyyy-MM-dd"),
                EDate = dateTimePicker_EDate.Value.ToString("yyyy-MM-dd"),
                WClass = comboBox_SearchHistory.SelectedValue.ToString();
            string string_SearchHistory =
                 "select wDate,WClass,WClassNm,EstiQty,ActQty,HuNum,HumanHr,startTime,RestTime,EstiTime,ActTime,ActHumanHr,WStatus,Memo from dbo.Px_WorkHr "
                + "where wDate between @BDate and @EDate ";
            if (WClass != "")
            {
                string_SearchHistory += "and WClass=@WClass";
                object[] obj_Param = { "@BDate", "DateTime", BDate, "@EDate", "DateTime", EDate, "@WClass", "NVarChar", WClass };
                dt_SerachHistory = sca.GetDataTable(Login_Server, string_SearchHistory, obj_Param, 0);
            }
            else
            {
                object[] obj_Param = { "@BDate", "DateTime", BDate, "@EDate", "DateTime", EDate, "@WClass", "NVarChar", WClass };
                dt_SerachHistory = sca.GetDataTable(Login_Server, string_SearchHistory, obj_Param, 0);
            }

            gridControl_SearchHistory.DataSource = dt_SerachHistory;
        }

        /// <summary>
        /// 儲存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_SearchHistory_Save_Click(object sender, EventArgs e)
        {
            int Modify_Count = 0;
            //是否有修改?
            DataTable Modify_Table = (DataTable)gridControl_SearchHistory.DataSource;
            if (Modify_Table == null)
            {
                MessageBox.Show("請先做查詢:)");
                return;
            }
            for (int i = Modify_Table.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr = Modify_Table.Rows[i];
                if (dr.RowState == DataRowState.Modified)
                    Modify_Count++;
            }

            if (Modify_Count > 0)
            {
                DialogResult dr1 = MessageBox.Show("確定修改資料??", "修改確認", MessageBoxButtons.YesNo);
                if (dr1 == System.Windows.Forms.DialogResult.No)
                    return;
                //儲存資料
                int SuccessUpdateCount = 0;
                string Update_string = "Update Px_WorkHr set Memo=@memo,Up_date=getdate(),UpUser=@LoginUserId where [WClass]=@WClass and [wDate]=@Date;Select SuccessCount=@@rowcount";
                string Result_string = "";
                for (int i = Modify_Table.Rows.Count - 1; i >= 0; i--)
                {
                    DataRow dr = Modify_Table.Rows[i];
                    if (dr.RowState == DataRowState.Modified)
                    {
                        object[] obj_Param = { "@memo", "NVarChar", dr["Memo"],
                                               "@WClass", "NVarChar", dr["WClass"],
                                               "@LoginUserId","NVarChar",fFormInfo.LoginUserId,
                                               "@Date","DateTime",dr["wDate"]};
                        DataTable test1 = sca.GetDataTable(Login_Server, Update_string, obj_Param, 0);
                        SuccessUpdateCount += Convert.ToInt32(test1.Rows[0]["SuccessCount"]);
                        if (Convert.ToInt32(test1.Rows[0]["SuccessCount"]) > 0)
                            Result_string += dr["wDate"].ToString() + " " + dr["WClassNm"].ToString() + " 成功更新備註\n";
                    }
                }
                if (SuccessUpdateCount > 0)
                {
                    Modify_Table.AcceptChanges();
                    MessageBox.Show(Result_string, "更新成功");
                }
            }
            else
            {
                MessageBox.Show("無資料被修改");
            }
        }

        /// <summary>
        /// 非今天的資料禁止修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        DateTime RowNowDate2 = DateTime.Now;
        private void gridView_SearchHistory_ShowingEditor(object sender, CancelEventArgs e)
        {
            DataRow dr = gridView_SearchHistory.GetFocusedDataRow();
            DateTime RowDate2 = DateTime.Parse(dr["wDate"].ToString());
            TimeSpan Difftime = new TimeSpan(RowNowDate2.Ticks - RowDate2.Ticks);
            if (Difftime.TotalHours > 48)
            {
                e.Cancel = true;
            }
        }

        /// <summary>
        /// 菲今天資料設定底色為灰色
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gridView_SearchHistory_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column == gridColumn37)
            {
                DateTime RowDate2 = DateTime.Parse(gridView_SearchHistory.GetRowCellValue(e.RowHandle, "wDate").ToString());
                TimeSpan Difftime = new TimeSpan(RowNowDate2.Ticks - RowDate2.Ticks);
                if (Difftime.TotalHours > 48)
                {
                    e.Appearance.BackColor = Color.LightGray;
                }
            }
        }

        /// <summary>
        /// 歷史紀錄匯出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            if (gridView_SearchHistory.RowCount <= 0)
            {
                MessageBox.Show("無內容，請重新查詢", "取消儲存");
                return;
            }
            SaveFileDialog OFD = new SaveFileDialog();
            OFD.InitialDirectory = "D:\\";
            OFD.RestoreDirectory = true;
            OFD.Title = "儲存檔案";
            OFD.DefaultExt = "xlsx";
            OFD.Filter = "Microsoft Office Excel 活頁簿 (*.xls)|*.xls";
            OFD.ShowDialog();
            if (OFD.FileName == "")
            {
                MessageBox.Show("未指定檔名，取消儲存Excel", "取消儲存");
                return;
            }
            gridView_SearchHistory.ExportToXlsx(OFD.FileName);
            MessageBox.Show("檔案匯出至 " + OFD.FileName + " 完畢");
        }

        #endregion

        #region 頁面_權限設定
        private void button_Authority_Query()
        {
            //先找在指定倉庫的權限
            int Authority_Seq = MyDictionaryFind(Table_SelStock, "Site_no", SelStock);
            if (Authority_Seq >= 0)
            {
                Authority = Convert.ToInt32(Table_SelStock.Rows[Authority_Seq]["Lv"]);
                if (Authority <= 2)
                {
                    button4.Enabled = true;
                }

                //Administrator SuperUser才能有權限設定功能
                if (Authority <= 1)
                {
                    //Reset
                    if (Table_Authority != null)
                    {
                        Table_Authority.RejectChanges();
                        Table_Authority.Clear();
                    }
                    Table_Authority_Del.RejectChanges();
                    Table_Authority_Del.Clear();

                    //查詢兩倉權限
                    string cmd_Authority = "Select Sn,Id,Name,Site_no,Lv,email from PxWorkHrRole where Site_no=@site_no";
                    object[] objTable = { "@site_no", "NChar", SelStock };
                    Table_Authority = sca.GetDataTable("EDI_DC", cmd_Authority, objTable, 0);
                    gridControl_Authority.DataSource = Table_Authority;

                }
            }

        }

        /// <summary>
        /// 新增權限設定明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemButtonEdit_Authority_Add_Click(object sender, EventArgs e)
        {
            DataRow Newdr = Table_Authority.NewRow();

            Newdr["Sn"] = -1;
            Newdr["Id"] = "";
            Newdr["Name"] = "";
            Newdr["Site_no"] = SelStock;
            Newdr["Lv"] = "5";

            Table_Authority.Rows.Add(Newdr);
        }

        /// <summary>
        /// 刪除權限設定明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemButtonEdit_Authority_Del_Click(object sender, EventArgs e)
        {
            //紀錄被刪除的SN
            int Del_sn = Convert.ToInt32(gridView_Authority.GetFocusedDataRow()["Sn"]);
            if (Del_sn >= 0)
            {
                Table_Authority_Del.Rows.Add(gridView_Authority.GetFocusedRowCellValue("Id"), gridView_Authority.GetFocusedRowCellValue("Site_no"));
            }
            gridView_Authority.DeleteRow(gridView_Authority.FocusedRowHandle);
        }

        /// <summary>
        /// 按下儲存按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Authority_Save_Click(object sender, EventArgs e)
        {
            string Table_Authority_InsertCommand = "Insert Into PxWorkHrRole values(@Id,@Name,@site_no,@Lv,@email)";
            string Table_Authority_UpdateCommand = "Update PxWorkHrRole set Lv=@Lv ,email=@email, Name=@Name where Id=@Id and Site_no=@site_no";
            string Table_Authority_DeleteCommand = "Delete from PxWorkHrRole where Id=@Id and Site_no=@site_no";
            try
            {
                foreach (DataRow dr in Table_Authority.Rows)
                {
                    //新增
                    if (dr.RowState == DataRowState.Added)
                    {
                        object[] objTable = { "@Id", "NChar", dr["Id"], "@Name", "NChar", dr["Name"], "@site_no", "NChar", dr["Site_no"], "@Lv", "Int", dr["Lv"], "@email", "NChar", dr["email"] };
                        sca.Update("EDI_DC", Table_Authority_InsertCommand, objTable, 0);
                    }
                    //更新
                    else if (dr.RowState == DataRowState.Modified)
                    {
                        object[] objTable = { "@Id", "NChar", dr["Id"], "@Name", "NChar", dr["Name"], "@site_no", "NChar", dr["Site_no"], "@Lv", "Int", dr["Lv"], "@email", "NChar", dr["email"] };
                        sca.Update("EDI_DC", Table_Authority_UpdateCommand, objTable, 0);
                    }
                }
                //刪除
                foreach (DataRow dr in Table_Authority_Del.Rows)
                {
                    object[] objTable = { "@Id", "NChar", dr["Id"], "@site_no", "NChar", dr["Site_no"] };
                    sca.Update("EDI_DC", Table_Authority_DeleteCommand, objTable, 0);
                }

                Table_Authority_Del.AcceptChanges();
                Table_Authority_Del.Clear();
                Table_Authority.AcceptChanges();
                button_Authority_Query();
            }
            catch (Exception e1)
            {
                Table_Authority_Del.RejectChanges();
                Table_Authority_Del.Clear();
                Table_Authority.RejectChanges();
                button_Authority_Query();
            }

            MessageBox.Show("更新完畢:)");
        }

        //禁止進入修改焦點
        private void gridView_Authority_ShowingEditor(object sender, CancelEventArgs e)
        {
            GridView view = sender as GridView;
            DataRow dr = gridView_Authority.GetFocusedDataRow();
            if (view.FocusedColumn.FieldName == "Id" && dr.RowState == DataRowState.Unchanged)
            {
                e.Cancel = true;
            }
        }

        //更改顏色
        private void gridView_Authority_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;
            if (e.Column == gridColumn_Authority_Id)
            {
                if (Convert.ToInt32(view.GetRowCellValue(e.RowHandle, "Sn")) > -1)
                {
                    e.Appearance.BackColor = Color.White;
                }
                else
                {
                    e.Appearance.BackColor = Color.FromArgb(255, 255, 192);
                }
            }
        }

        //顯示/不顯示 權限說明
        private void btn_Authority_Help_Click(object sender, EventArgs e)
        {
            panel1.Visible = !panel1.Visible;
        }

        #endregion

        #region 頁面_作業完成回填
        /// <summary>
        /// 取得作業完成回填的清單
        /// </summary>
        private void CrtWorkData()
        {
            try
            {
                dtWk = dtWorkData();
                gv_WkList.DataSource = dtWk;
                Light();
            }
            catch (Exception)
            { }
        }

        //即時查詢
        private void button1_Click(object sender, EventArgs e)
        {
            CrtWorkData();
        }

        //gv_wklist選擇cell
        private void gv_WkList_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int intFinBtn = 1;
            bool blSubmit = false;

            if (e.ColumnIndex == intFinBtn)
            {
                //string gg=gv_WkList.CurrentRow.Cells[1].Value.ToString();
                //string strStatus= ((DataGridView)sender).Rows[e.RowIndex].Cells[1].Value.ToString();
                string strClass = ((DataGridView)sender).Rows[e.RowIndex].Cells[3].Value.ToString();
                string strStatus = ((DataGridView)sender).Rows[e.RowIndex].Cells[6].Value.ToString();
                //string strDate = ((DataGridView)sender).Rows[e.RowIndex].Cells[4].Value.ToString();
                //gv_WkList.CurrentRow.Cells[1].Value = "dd";
                //string a = gv_WkList.CurrentRow.Cells[2].Value.ToString();
                //string b = gv_WkList.CurrentRow.Cells[4].Value.ToString();
                if (strStatus == "0")
                {
                    blSubmit = UpStatus(1, strClass);

                }
                else
                {
                    blSubmit = UpStatus(0, strClass);

                }
                if (blSubmit)
                {

                    CrtWorkData();

                    CkWorkFin();

                }
            }
        }

        #region 送信
        /// <summary>
        /// 檢核工作完成
        /// </summary>
        private void CkWorkFin()
        {

            DataView view = new DataView();
            DataTable dtCk = new DataTable();
            view.Table = dtWk;
            view.RowFilter = "isnull(WStatus,0) = '0'";
            dtCk = view.ToTable();
            if (dtCk.Rows.Count == 0)
            {
                Application.DoEvents();
                label9.Text = "開始發信勿關閉程式";
                timer1.Start();
                DataTable dtRp = new DataTable();
                DataTable dtMail = new DataTable();
                dtRp = dtReport();
                dtMail = dtUserMail2();
                SendRep(dtRp, dtMail);
                timer1.Stop();
                label9.Text = "";
            }
            MessageBox.Show("執行完成");
        }
        /// <summary>
        /// 啟動系統派信
        /// </summary>
        /// <param name="dtRp"></param>
        /// <param name="ToMail"></param>
        private void SendRep(DataTable dtRp, DataTable ToMail)
        {
            SMM SM = new SMM();
            int MailNum = ToMail.Rows.Count;
            string WorkDate = DateTime.Now.ToString("yyyy/MM/dd");
            string[] arrMail = new string[MailNum];
            string[] arrName = new string[MailNum];
            string[] CC = new string[1];
            string[] CCName = new string[1];
            for (int i = 0; i < MailNum; i++)
            {
                arrMail[i] = ToMail.Rows[i]["Email"].ToString();
                arrName[i] = ToMail.Rows[i]["Name"].ToString();
            }
            CC[0] = "Jack_cho@pxmart.com.tw";
            CCName[0] = "Jack_cho";

            string strMsg = string.Empty;
            strMsg += "敬會 長官 您好:";
            strMsg += "<p/> 下表為今日 " + DCName + WorkDate + "營運作業量 <br/>";
            SM.TbMail(
                "mail.pxmart.com.tw",
                25,
                false,
                "WebManager@pxmart.com.tw",
                "WM23578000",
                DCName + " 物流作業量自動發信",
                "WebManager@pxmart.com.tw",
                DCName + " 物流作業量自動發信",
                arrMail,
                arrName,
                CC,
                CCName,
                strMsg,
                DCName + "營運作業量",
                dtRp);
        }
        /// <summary>
        /// 啟動系統補發信
        /// </summary>
        /// <param name="dtRp"></param>
        /// <param name="ToMail"></param>
        private void SendRep2(DataTable dtRp, DataTable ToMail)
        {
            SMM SM = new SMM();
            int MailNum = ToMail.Rows.Count;
            string WorkDate = DateTime.Now.ToString("yyyy/MM/dd");
            string[] arrMail = new string[MailNum];
            string[] arrName = new string[MailNum];
            string[] CC = new string[1];
            string[] CCName = new string[1];
            for (int i = 0; i < MailNum; i++)
            {
                arrMail[i] = ToMail.Rows[i]["Email"].ToString();
                arrName[i] = ToMail.Rows[i]["Name"].ToString();
            }
            CC[0] = "Jack_cho@pxmart.com.tw";
            CCName[0] = "Jack_cho";

            string strMsg = string.Empty;
            strMsg += "敬會 長官 您好:";
            strMsg += "<p/>物流作業均已完成請上人效確認 : 下表為今日 " + DCName + WorkDate + "營運作業量 <br/>";
            SM.TbMail(
                "mail.pxmart.com.tw",
                25,
                false,
                "WebManager@pxmart.com.tw",
                "WM23578000", DCName + " 物流作業均已完成請上人效確認",
                "WebManager@pxmart.com.tw",
                DCName + " 物流作業均已完成請上人效確認",
                arrMail,
                arrName,
                CC,
                CCName,
                strMsg,
                DCName + "營運作業量",
                dtRp);
        }
        /// <summary>
        /// 取得人效作業量
        /// </summary>
        /// <returns></returns>
        private DataTable dtReport()
        {
            DataTable dt = new DataTable();

            try
            {
                dt = sca.GetDataTable(Login_Server, String.Format("sp_PxImdWorkmail ", DateTime.Now.ToString("yyyy/MM/dd")));
            }
            catch (Exception)
            { }
            finally
            {

            }
            return dt;
        }
        /// <summary>
        /// 取得收件者名單
        /// </summary>
        /// <returns></returns>
        private DataTable dtUserMail()
        {
            DataTable dt = new DataTable();

            try
            {

                dt = sca.GetDataTable(Login_Server, "select ID,Name,Email from PxWorkHrRole where  Lv<=1 and isnull(email,'')<>'' and site_no='{0}' ");
            }
            catch (Exception)
            {

            }
            finally
            {

            }
            return dt;
        }
        /// <summary>
        /// 補發:取得收件者名單
        /// </summary>
        /// <returns></returns>
        private DataTable dtUserMail2()
        {
            DataTable dt = new DataTable();

            try
            {

                dt = sca.GetDataTable("EDI_DC", String.Format("select ID,Name,Email from PxWorkHrRole where Lv<>0 and Lv<=3 and isnull(email,'')<>'' and site_no='{0}' ", SelStock.ToString()));
            }
            catch (Exception)
            {

            }
            finally
            {

            }
            return dt;
        }

        //確認並發信
        private void button6_Click(object sender, EventArgs e)
        {
            timer1.Start();
            label9.Text = "開始發信勿關閉程式";
            DataTable dtRp = new DataTable();
            DataTable dtMail = new DataTable();
            dtRp = dtReport();
            dtMail = dtUserMail();
            SendRep(dtRp, dtMail);
            MessageBox.Show("發信完成");
            timer1.Stop();
            label9.Text = "";
        }
        //補發作業完成通知信
        private void button4_Click(object sender, EventArgs e)
        {
            label9.Text = "開始發信勿關閉程式";
            timer1.Start();
            DataTable dtRp = new DataTable();
            DataTable dtMail = new DataTable();
            dtRp = dtReport();
            dtMail = dtUserMail2();
            SendRep2(dtRp, dtMail);
            MessageBox.Show("發信完成");
            timer1.Stop();
            label9.Text = "";
        }
        #endregion

        #endregion

        #region 頁面_RF條碼維護
        /// <summary>
        /// 查詢RF條碼設定
        /// </summary>
        private void button_work_Query()
        {
            switch (Login_Server)
            {
                case "pxwms_n": SelStock = "觀音倉"; break;
                case "pxwms_s": SelStock = "岡山倉"; break;
                case "pxwms_c_HumanHr": SelStock = "梧棲倉"; break;
                default: SelStock = "觀音倉"; break;
            }
            //先找在指定倉庫的權限
            int Authority_Seq = MyDictionaryFind(Table_SelStock, "Site_no", SelStock);
            if (Authority_Seq >= 0)
            {
                int Authority = Convert.ToInt32(Table_SelStock.Rows[Authority_Seq]["Lv"]);

                //Administrator SuperUser才能有權限設定功能
                if (Authority <= 5)
                {
                    //Reset
                    if (Table_RFbarcode != null)
                    {
                        Table_RFbarcode.RejectChanges();
                        Table_RFbarcode.Clear();
                    }
                    Table_RFbarcode_Del.RejectChanges();
                    Table_RFbarcode_Del.Clear();

                    //查詢兩倉權限

                }
            }
            string cmd_Authority = "select DCTYPE ,GROUPNAME ,ITEMNAME ,barcode,isnull(wkclass,'') as wkclass,barcode as OriginalBarcode from workitem_jack where DCTYPE=@site_no";
            object[] objTable = { "@site_no", "NChar", SelStock };
            Table_RFbarcode = sca.GetDataTable(Login_Server, cmd_Authority, objTable, 0);
            gridControl_RFbarcode.DataSource = Table_RFbarcode;


        }

        /// <summary>
        /// 組別的下拉式選單
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemComboBox_Groupname_SelectedIndexChanged(object sender, EventArgs e)
        {
            string value = (sender as DevExpress.XtraEditors.ComboBoxEdit).Text;

            ////1.获取下拉框选中值
            //Item item = (Item)(sender as ComboBoxEdit).SelectedItem;
            //string text = item.Name.ToString();
            //string value = item.Value;
            //2.获取gridview选中的行
            GridView myView = (gridControl_RFbarcode.MainView as GridView);
            int dataIndex = myView.GetDataSourceRowIndex(myView.FocusedRowHandle);
            //3.保存选中值到datatable
            Table_RFbarcode.Rows[dataIndex]["groupname"] = value;
        }

        /// <summary>
        /// 新增權限設定明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemButtonEdit_Authority_Add2_Click(object sender, EventArgs e)
        {
            DataRow Newdr = Table_RFbarcode.NewRow();

            Newdr["DCTYPE"] = SelStock;
            Newdr["GROUPNAME"] = "";
            Newdr["ITEMNAME"] = "";
            Newdr["barcode"] = "";
            Newdr["wkclass"] = "";

            Table_RFbarcode.Rows.Add(Newdr);
        }

        /// <summary>
        /// 刪除權限設定明細
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void repositoryItemButtonEdit_Authority_Del2_Click(object sender, EventArgs e)
        {
            //紀錄被刪除的SN
            string Del_sn = gridView_RFbarcode.GetFocusedDataRow()["barcode"].ToString();
            Table_RFbarcode_Del.Rows.Add(gridView_RFbarcode.GetFocusedRowCellValue("barcode"));
            gridView_RFbarcode.DeleteRow(gridView_RFbarcode.FocusedRowHandle);
        }

        /// <summary>
        /// 按下儲存按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_RFbarcode_Save_Click(object sender, EventArgs e)
        {
            string Table_Authority_InsertCommand = "Insert Into workitem_jack values(@DCtype,@groupname,@itemname,@barcode,@wkclass)";
            string Table_Authority_UpdateCommand = "Update workitem_jack set groupname=@groupname, itemname=@itemname, wkclass=@wkclass, barcode=@barcode where barcode=@OriginalBarcode";
            string Table_Authority_DeleteCommand = "Delete from workitem_jack where barcode=@barcode";
            try
            {
                foreach (DataRow dr in Table_RFbarcode.Rows)
                {
                    if (dr.RowState != DataRowState.Deleted)
                    {
                        object[] objTable = { "@DCType", "NVarChar", dr["DCTYPE"].ToString(),
                                          "@GROUPNAME", "NVarChar", dr["GROUPNAME"].ToString(),
                                          "@ITEMNAME", "NVarChar", dr["ITEMNAME"].ToString(),
                                          "@barcode", "NVarChar", dr["barcode"].ToString(),
                                          "@wkclass", "NVarChar", dr["wkclass"].ToString(),
                                            "@OriginalBarcode", "NVarChar", dr["OriginalBarcode"].ToString()};

                        //新增
                        if (dr.RowState == DataRowState.Added)
                        {
                            sca.Update(Login_Server, Table_Authority_InsertCommand, objTable, 0);
                        }
                        //更新
                        else if (dr.RowState == DataRowState.Modified)
                        {
                            sca.Update(Login_Server, Table_Authority_UpdateCommand, objTable, 0);
                        }
                    }
                }
                //刪除
                foreach (DataRow dr in Table_RFbarcode_Del.Rows)
                {
                    object[] objTable = { "@barcode", "NVarChar", dr["barcode"] };
                    sca.Update(Login_Server, Table_Authority_DeleteCommand, objTable, 0);
                }

                Table_RFbarcode_Del.AcceptChanges();
                Table_RFbarcode_Del.Clear();
                Table_RFbarcode.AcceptChanges();
                button_work_Query();
            }
            catch (Exception)
            {
                Table_RFbarcode_Del.RejectChanges();
                Table_RFbarcode_Del.Clear();
                Table_RFbarcode.RejectChanges();
                button_work_Query();
            }

            MessageBox.Show("更新完畢:)");
        }

        #endregion

        #region 頁面_組別維護
        /// <summary>
        /// 查詢組別
        /// </summary>
        private void button_workreferce_Query()
        {
            string strwork_referce = "select * from work_referce";
            object[] obj_work_referce = { "@userID", "NChar", label_EMPID.Text };
            Table_workreferce = sca.GetDataTable(Login_Server, strwork_referce, obj_work_referce, 0);
            gridControl_workreferce.DataSource = Table_workreferce;

            repositoryItemComboBox_Groupname.Items.Clear();
            foreach (DataRow dr_work_referce in Table_workreferce.Rows)
            {
                repositoryItemComboBox_Groupname.Items.Add(dr_work_referce["groupname"].ToString());
            }
        }

        /// <summary>
        /// 組別儲存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_workreferce_Save_Click(object sender, EventArgs e)
        {
            string Table_Authority_UpdateCommand = "Update work_referce set groupname=@groupname where groupid=@groupid";

            try
            {
                foreach (DataRow dr in Table_workreferce.Rows)
                {
                    //更新
                    if (dr.RowState == DataRowState.Modified)
                    {
                        object[] objTable = {"@GROUPNAME", "NVarChar", dr["groupname"].ToString(),
                                          "@groupid", "NVarChar", dr["groupid"].ToString()};

                        sca.Update(Login_Server, Table_Authority_UpdateCommand, objTable, 0);
                    }
                }

                Table_workreferce.AcceptChanges();
                button_workreferce_Query();
            }
            catch (Exception)
            {
                Table_workreferce.RejectChanges();
                button_workreferce_Query();
            }

            MessageBox.Show("更新完畢:)");
        }
        #endregion

        #region 頁面_刷卡紀錄
        /// <summary>
        /// 查詢刷卡紀錄
        /// </summary>
        private void btn_RecordList_Query()
        {
            string uri;
            switch (Login_Server)
            {
                case "pxwms_n":
                    uri = "WMSS0353"; break;
                case "pxwms_s":
                    uri = "WMSS0352"; break;
                case "pxwms_c_HumanHr":
                    uri = "WMSS0354"; break;
                default: uri = ""; break;
            }
            axSamrtViewImpl11.LinkQuery("192.168.110.70", "kota", "1234", "DRP", "DRP_SEARCH", uri, false, "PXMART", "");
        }

        /// <summary>
        /// 查詢控管報表
        /// </summary>
        private void btn_ControlList_Query()
        {
            string uri;
            switch (Login_Server)
            {
                case "pxwms_n":
                    uri = "WMSS0343"; break;
                case "pxwms_s":
                    uri = "WMSS0342"; break;
                case "pxwms_c_HumanHr":
                    uri = "WMSS0344"; break;
                default: uri = ""; break;
            }
            axSamrtViewImpl12.LinkQuery("192.168.110.70", "kota", "1234", "DRP", "DRP_SEARCH", uri, false, "PXMART", "");

        }
        #endregion

        #region Form_View
        /// <summary>
        /// 分頁切換
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button_Settings_Save.Enabled = false;

            switch (tabControl1.SelectedIndex)
            {
                case 0: //設定
                    button_Settings_Query();
                    ((DataSet)gridControl_Settings.DataSource).RejectChanges();
                    button_Settings_Save.Enabled = true;
                    gridControl_Settings.Focus(); break;
                case 1: //即時查詢
                    break;
                case 2: //歷史查詢
                    break;
                case 3: //權限設定
                    button_Authority_Query();
                    break;
                case 4: //權限設定
                    CrtWorkData();
                    break;
                case 5: //RF條碼
                    button_work_Query();
                    break;
                case 6: //組別維護
                    button_workreferce_Query();
                    break;
                case 7: //SQ報表
                    tabControl2_SelectedIndexChanged(tabControl2, e);
                    break;
            }
        }

        /// <summary>
        /// 報表:分頁切換
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl2.SelectedIndex)
            {
                case 0: //刷卡紀錄
                    btn_RecordList_Query(); break;
                case 1: //控管報表
                    btn_ControlList_Query(); break;
                default: break;
            }
        }

        /// <summary>
        /// 表單按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = MessageBox.Show("確定離開  " + this.Text + "??", "離開確認", MessageBoxButtons.YesNo);
                if (this.DialogResult == DialogResult.Yes)
                {
                    this.Close();
                }
                else
                {
                    this.DialogResult = DialogResult.None;
                }
            }
        }


        #region 自定義Format格式
        public class XXXXFormatter : IFormatProvider, ICustomFormatter
        {
            public XXXXFormatter() { }

            public object GetFormat(System.Type type)
            {
                return this;
            }

            public string Format(string formatString, object arg, IFormatProvider formatProvider)
            {
                switch (arg.ToString())
                {
                    case "0": formatString = "準時"; break;
                    case "1": formatString = "提早"; break;
                    case "2": formatString = "延遲"; break;
                    default: formatString = ""; break;
                }
                return formatString;
                //注意这里我们返回的是string值.
                //formatString 就是你自己定义的format方法,需要在这里实现判断
                //arg 就是这个单元格的值, 或者说是 绑定datasource后的字段值
                //formatProvider 这个就可以不用管他的,说实话我也没有搞清楚怎么用,知道的朋友补充一下吧!
            }
        }
        #endregion
        #endregion

        #region Form_Control
        /// <summary>
        /// 切換倉別
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox_Stock_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_Stock.SelectedIndex < 0)
                return;
            string Search_WClass = ((Item)comboBox_Stock.SelectedItem).Value;
            Login_Server = Search_WClass;

            switch (comboBox_Stock.SelectedIndex)
            {
                case 0: txb_http.Text = "http://172.20.210.10/rf_box/funList_07.aspx"; break;
                case 1: txb_http.Text = "http://172.20.236.10/rf_box/funList_07.aspx"; break;
                case 2: txb_http.Text = "http://172.20.250.16/rf_box/rf_box/funList_07.aspx"; break;
            }
            switch (Login_Server)
            {
                case "pxwms_n": SelStock = "DC01"; DCName = "觀音倉"; break;
                case "pxwms_s": SelStock = "DC02"; DCName = "岡山倉"; break;
                case "pxwms_c_HumanHr": SelStock = "DC03"; DCName = "梧棲倉"; break;
                default: SelStock = "DC01"; DCName = "觀音倉"; break;
            }

            tabControl1_SelectedIndexChanged(sender, e);
        }

        /// <summary>
        /// IO取得完成狀態
        /// </summary>
        /// <returns></returns>
        private DataTable dtWorkData()
        {
            DataTable dt = new DataTable();


            dt = sca.GetDataTable(Login_Server, "Select WClass,WClassNm,WDate,isnull(WStatus,0) as WStatus,acttime as FinishTime,upuser as 最後更新人員,up_date as 更新時間 From Px_WorkHr  where WDate=convert(varchar(10),getdate(),121) and isnull(DelFlg,'') <>'Y' ");
            return dt;
        }

        /// <summary>
        /// 判別紅綠燈
        /// </summary>
        private void Light()
        {
            string GreLight = Application.StartupPath + "\\media-record-3_2.png";
            string RedLight = Application.StartupPath + "\\media-record-3.png";
            DataTable dt = new DataTable();
            dt = (DataTable)gv_WkList.DataSource;
            int gvCount = gv_WkList.Rows.Count;
            gv_WkList[0, 0].Style.BackColor = Color.Green;
            sk = gvCount;
            for (int i = 0; i < gvCount; i++)
            {
                if (gv_WkList[6, i].Value.ToString().Trim() == "1")
                {
                    sk = sk - 1;
                }

            }


            if (sk == 0 && Authority <= 2)
            {
                gv_WkList.Columns["Column1"].Visible = true;

                for (int i = 0; i < gvCount; i++)
                {
                    if (gv_WkList[6, i].Value.ToString().Trim() == "0")
                    {

                        gv_WkList[0, i].Style.BackColor = Color.Red;
                        gv_WkList[1, i].Value = "完成";

                    }
                    else
                    {
                        gv_WkList[0, i].Style.BackColor = Color.Green;
                        gv_WkList[1, i].Value = "解鎖";
                    }
                    gv_WkList[2, i].Value = "否決";
                }
                button6.Visible = true;
            }
            else if (sk == 0 && Authority > 2)
            {
                button6.Visible = false;
                gv_WkList.Columns["col_Finish"].Visible = false;
            }
            else
            {
                button6.Visible = false;
                gv_WkList.Columns["Column1"].Visible = false;
                gv_WkList.Columns["col_Finish"].Visible = true;
                for (int i = 0; i < gvCount; i++)
                {
                    if (gv_WkList[6, i].Value.ToString().Trim() == "0")
                    {

                        gv_WkList[0, i].Style.BackColor = Color.Red;
                        gv_WkList[1, i].Value = "完成";

                    }
                    else
                    {
                        gv_WkList[0, i].Style.BackColor = Color.Green;
                        gv_WkList[1, i].Value = "取消";

                    }
                }
            }
        }

        /// <summary>
        /// IO更新狀態
        /// </summary>
        /// <param name="intStatus"></param>
        /// <param name="wClass"></param>
        /// <returns></returns>
        private bool UpStatus(int intStatus, string wClass)
        {
            bool blStatus = false;
            // string con_str = Con("pxwms");
            // SqlConnection SqlConn = new SqlConnection(con_str);
            object strFinTime;
            if (intStatus == 1)
            {
                strFinTime = DateTime.Now;
            }
            else
            {
                strFinTime = DBNull.Value;
            }
            try
            {
                string Sql_cmd = "";
                if (intStatus == 1)
                {
                    Sql_cmd = string.Format("Update Px_WorkHr  set acttime=getdate() , WStatus='1' ,Up_date=getdate(),UpUser='{1}' Where WClass='{0}' and WDate=convert(varchar(10),getdate(),121)  ", wClass, fFormInfo.LoginUserId);

                }
                else
                {
                    Sql_cmd = string.Format("Update Px_WorkHr  set acttime=null , WStatus='0' ,Up_date=getdate(),UpUser='{1}' Where WClass='{0}' and WDate=convert(varchar(10),getdate(),121) ", wClass, fFormInfo.LoginUserId);
                }

                sca.Update(Login_Server, Sql_cmd);
                if (intStatus == 1)
                {
                    Sql_cmd = string.Format("update  Darren_RFWorkRecord_Agg set wETime = getdate(),wTime = datediff(mi,wSTime,getdate()) ", wClass, fFormInfo.LoginUserId);
                    Sql_cmd = Sql_cmd + string.Format(" where wETime is null and CONVERT(char(8),wSTime,112)=CONVERT(char(8),GETDATE(),112) and ", wClass, fFormInfo.LoginUserId);
                    Sql_cmd = Sql_cmd + string.Format(" REPLACE(wZone,' ','') in (select dctype+GROUPNAME+ITEMNAME from workitem_jack where wkclass='{0}')", wClass, fFormInfo.LoginUserId);
                    if (MessageBox.Show("是否一並將人員簽退", "是否一並將人員簽退", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) == DialogResult.Yes)
                    {

                        sca.Update(Login_Server, Sql_cmd);
                    }
                }
                blStatus = true;

            }
            catch (Exception)
            {
                blStatus = false;
            }
            finally
            {

            }
            return blStatus;
        }

        /// <summary>
        /// 字典搜尋
        /// </summary>
        /// <param name="dt_Source">來源Table</param>
        /// <param name="ColumnName">欲查詢欄位</param>
        /// <param name="SearchString">查詢值</param>
        /// <returns>行數代號</returns>
        private int MyDictionaryFind(DataTable dt_Source, string ColumnName, string SearchString)
        {
            int Search_Index = -1;
            SearchString = SearchString.Trim().ToUpper();

            //建立Dictionary
            Dictionary<string, int> MyDictionary = new Dictionary<string, int>();
            int i = 0;
            foreach (DataRow dr in dt_Source.Rows)
            {
                MyDictionary.Add(dr[ColumnName].ToString().Trim().ToUpper(), i);
                i++;
            }
            //查詢
            if (MyDictionary.ContainsKey(SearchString))
            {
                Search_Index = MyDictionary[SearchString];
            }
            return Search_Index;
        }

        #endregion

        //label9變色?
        private void timer1_Tick(object sender, EventArgs e)
        {
            int t1 = 0;
            if (t1 == 0)
            {

                label9.BackColor = Color.DarkOrange;
                t1 = 1;
            }
            else
            {

                BackColor = Color.Blue;
                t1 = 0;
            }
        }
    }

    //下拉式選單內容
    public class Item
    {
        public string Name;
        public string Value;
        public Item(string name, string value)
        {
            Name = name; Value = value;
        }
        public override string ToString()
        {
            // Generates the text shown in the combo box
            return Name;
        }
    }





}
