/*************************************************
 * 株式会社吉伸
 * タグラベル出力システム
 * -----------------------------------------------
 * 機能名：三陽商会 ＥＤＩ Rainbow連携画面
 * -----------------------------------------------
 * 更新履歴
 * 2020.02.01   TCS saita        新規作成
 * 2020.02.12   TCS sato         依頼明細票2019-049,051
 *
 * 
 *************************************************/
using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace タグラベル出力システム
{
    public partial class frmSNYEDIRainbowCSV : Form
    {
        static SqlCommand sqlSel;
        static SqlCommand sqlSelCsv;
        static SqlCommand sqlIdentity;
        static SqlConnection sqlCon;
        static SqlConnection sqlConIdentity;
        private DataGridViewTextBoxColumn DGVText;
        private DataGridViewCheckBoxColumn DGVCheckBox;
        private DataGridViewRow row;

        private string sDisplay_flg = clConst.cOff;
        private int identityNo = 0;
        private const int rowCnt = 6;
        private const int colCnt = 8;

        public frmSNYEDIOrder order;

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public frmSNYEDIRainbowCSV()
        {
            InitializeComponent();
        }
        #endregion

        #region 画面ロード
        /// <summary>
        /// 画面ロード
        /// </summary>
        private void frmSNYEDIRainbowCSV_Load(object sender, EventArgs e)
        {
            // DB接続文字列を生成
            clConnect.getConnect();
            sqlCon = new System.Data.SqlClient.SqlConnection(clConnect.sCn);
            sqlConIdentity = new System.Data.SqlClient.SqlConnection(clConnect.sCn);

            //画面初期化
            if (this.Owner == null)
            {
                Form_Initialize();
            }
            else
            {
                btnCSV.Focus();
            }
        }
        #endregion

        #region 画面初期化
        /// <summary>
        /// 画面初期化
        /// </summary>
        private void Form_Initialize()
        {
            txtRecieveCnt.Enabled = true;
            txtRecieveYmd.Enabled = true;
            txtShippingYmd.Enabled = true;
            txtOrderNo.Enabled = true;
            txtInstNo.Enabled = true;

            txtRecieveCnt.Text = string.Empty;
            txtRecieveYmd.Text = string.Empty;
            txtShippingYmd.Text = string.Empty;
            txtOrderNo.Text = string.Empty;
            txtInstNo.Text = string.Empty;
            lblResult.Text = string.Empty;
            txtBrName.Text = string.Empty;

            sDisplay_flg = clConst.cOff;

            dgvEDIOrderList_Initialize();

            rdoAll.Enabled = true;
            rdoNuno.Enabled = true;
            rdoSagefuda.Enabled = true;
            rdoAll.Checked = true;

            txtBrName.Enabled = true;

            this.KeyPreview = true;
        }
        #endregion

        #region データグリッドビュー初期化
        /// <summary>
        /// データグリッドビュー初期化
        /// </summary>
        private void dgvEDIOrderList_Initialize()
        {
            dgvEDIOrderList.Columns.Clear();
            dgvEDIOrderList.AllowUserToResizeRows = false; // 高さ変更不可
            // ヘッダー設定（中央に配置）
            dgvEDIOrderList.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //対象
            this.DGVCheckBox = new DataGridViewCheckBoxColumn();
            this.DGVCheckBox.Name = "chkdetail";
            this.DGVCheckBox.HeaderText = "";
            this.DGVCheckBox.AutoSizeMode = DataGridViewAutoSizeColumnMode.NotSet;
            this.DGVCheckBox.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.DGVCheckBox.Width = 50;
            this.DGVCheckBox.ReadOnly = false;
            this.DGVCheckBox.Visible = true;
            this.DGVCheckBox.FalseValue = "0";
            this.DGVCheckBox.TrueValue = "1";
            this.DGVCheckBox.Frozen = true;
            this.dgvEDIOrderList.Columns.Add(DGVCheckBox);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "inst_no";
            this.DGVText.HeaderText = "指図書番号";
            this.DGVText.Width = 100;
            this.DGVText.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = true;
            this.DGVText.SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "order_no";
            this.DGVText.HeaderText = "伝票NO";
            this.DGVText.Width = 150;
            this.DGVText.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = true;
            this.DGVText.SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "order_kbn";
            this.DGVText.Width = 0;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "br_name";
            this.DGVText.HeaderText = "ブランド";
            this.DGVText.Width = 400;
            this.DGVText.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = true;
            this.DGVText.SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "shipping_ymd";
            this.DGVText.HeaderText = "出荷日";
            this.DGVText.Width = 110;
            this.DGVText.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = true;
            this.DGVText.SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "received_ymd";
            this.DGVText.HeaderText = "受信日";
            this.DGVText.Width = 110;
            this.DGVText.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = true;
            this.DGVText.SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "received_number";
            this.DGVText.HeaderText = "受信回数";
            this.DGVText.Width = 95;
            this.DGVText.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.DGVText.ReadOnly = true;
            this.DGVText.Visible = true;
            this.DGVText.SortMode = DataGridViewColumnSortMode.NotSortable;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag1_cd";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag2_cd";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag3_cd";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag4_cd";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag5_cd";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag6_cd";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag1_cnt";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag2_cnt";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag3_cnt";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag4_cnt";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag5_cnt";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "tag6_cnt";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "identity_flg1";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "identity_flg2";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "identity_flg3";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "identity_flg4";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "identity_flg5";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);

            this.DGVText = new DataGridViewTextBoxColumn();
            this.DGVText.Name = "identity_flg6";
            this.DGVText.Width = 0;
            this.DGVText.Visible = false;
            this.dgvEDIOrderList.Columns.Add(DGVText);
        }
        #endregion

        #region 全選択/全解除チェックボックスイベント
        /// <summary>
        /// チェック全選択全解除変更時処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkAll_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAll.Checked)
            {
                // グリッドチェックボックスを全選択
                for (var i = 0; i < dgvEDIOrderList.RowCount; i++)
                {
                    dgvEDIOrderList["chkdetail", i].Value = clConst.cOn;
                }
            }
            else
            {
                // グリッドチェックボックスを全解除
                for (var i = 0; i < dgvEDIOrderList.RowCount; i++)
                {
                    dgvEDIOrderList["chkdetail", i].Value = clConst.cOff;
                }
            }
            dgvEDIOrderList.Refresh();
        }
        #endregion

        #region 表示ボタン押下時
        /// <summary>
        /// 表示ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void btnDisplay_Click(object sender, EventArgs e)
        {
            bool chkWhere = false;

            string sSQL;
            string sWhere = string.Empty;
            string sRecieveCnt;
            string sRecieveYmd;
            string sShippingYmd;
            string sInstNo;
            string sOrderNo;
            string sOrderKbn;
            string sBrName;

            dgvEDIOrderList_Initialize();

            if (!string.IsNullOrEmpty(txtRecieveYmd.Text))
            {
                if (string.IsNullOrEmpty(txtRecieveCnt.Text))
                {
                    MessageBox.Show("受信回数を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(txtRecieveCnt.Text))
                {
                    MessageBox.Show("受信日を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (!string.IsNullOrEmpty(txtInstNo.Text))
            {
                if (string.IsNullOrEmpty(txtOrderNo.Text))
                {
                    MessageBox.Show("伝票NOを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(txtOrderNo.Text))
                {
                    MessageBox.Show("指図書番号を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            try
            {
                // DB 接続
                clConnect.getConnect();
                sqlCon = new System.Data.SqlClient.SqlConnection(clConnect.sCn);
                sqlCon.Open();
                sqlSel = sqlCon.CreateCommand();

                //t_SNY_EDI_ORDERを検索
                sSQL  = "select EDI_STATUS.inst_no as inst_no,";
                sSQL += "       EDI_STATUS.order_no as order_no,";
                sSQL += "       EDI_ORDER.order_kbn as order_kbn,";
                sSQL += "       EDI_ORDER.br_name as br_name,";
                sSQL += "       EDI_STATUS.shipping_ymd as shipping_ymd,";
                sSQL += "       EDI_ORDER.received_ymd as received_ymd,";
                sSQL += "       EDI_ORDER.received_number as received_number,";
                sSQL += "       EDI_ORDER.tag1_cd,";
                sSQL += "       EDI_ORDER.tag1_cnt,";
                sSQL += "       isnull(CODE1.identity_flg, '') as identity_flg1,";
                sSQL += "       EDI_ORDER.tag2_cd,";
                sSQL += "       EDI_ORDER.tag2_cnt,";
                sSQL += "       isnull(CODE2.identity_flg, '') as identity_flg2,";
                sSQL += "       EDI_ORDER.tag3_cd,";
                sSQL += "       EDI_ORDER.tag3_cnt,";
                sSQL += "       isnull(CODE3.identity_flg, '') as identity_flg3,";
                sSQL += "       EDI_ORDER.tag4_cd,";
                sSQL += "       EDI_ORDER.tag4_cnt,";
                sSQL += "       isnull(CODE4.identity_flg, '') as identity_flg4,";
                sSQL += "       EDI_ORDER.tag5_cd,";
                sSQL += "       EDI_ORDER.tag5_cnt,";
                sSQL += "       isnull(CODE5.identity_flg, '') as identity_flg5,";
                sSQL += "       EDI_ORDER.tag6_cd,";
                sSQL += "       EDI_ORDER.tag6_cnt,";
                sSQL += "       isnull(CODE6.identity_flg, '') as identity_flg6";
                sSQL += "  from (select inst_no,";
                sSQL += "               order_no,";
                sSQL += "               shipping_ymd";
                sSQL += "          from t_SNY_EDI_STATUS";
                sSQL += "        group by inst_no, order_no, shipping_ymd";
                sSQL += "       ) EDI_STATUS";
                sSQL += "       inner join t_SNY_EDI_ORDER as EDI_ORDER on EDI_STATUS.inst_no = EDI_ORDER.inst_no";
                sSQL += "                                              and EDI_STATUS.order_no = EDI_ORDER.order_no";
                sSQL += "                                              and EDI_ORDER.order_kbn != 'F'";    //附属以外を対象
                sSQL += "                                              and EDI_ORDER.delete_flg = '0'";
                sSQL += "       left outer join m_SNY_CODE as CODE1 on EDI_ORDER.tag1_cd = CODE1.code_no";
                sSQL += "                                     and CODE1.code_kbn = 'S'";
                sSQL += "                                     and CODE1.code_sub_kbn = 'B'";
                sSQL += "       left outer join m_SNY_CODE as CODE2 on EDI_ORDER.tag2_cd = CODE2.code_no";
                sSQL += "                                     and CODE2.code_kbn = 'S'";
                sSQL += "                                     and CODE2.code_sub_kbn = 'B'";
                sSQL += "       left outer join m_SNY_CODE as CODE3 on EDI_ORDER.tag3_cd = CODE3.code_no";
                sSQL += "                                     and CODE3.code_kbn = 'S'";
                sSQL += "                                     and CODE3.code_sub_kbn = 'B'";
                sSQL += "       left outer join m_SNY_CODE as CODE4 on EDI_ORDER.tag4_cd = CODE4.code_no";
                sSQL += "                                     and CODE4.code_kbn = 'S'";
                sSQL += "                                     and CODE4.code_sub_kbn = 'B'";
                sSQL += "       left outer join m_SNY_CODE as CODE5 on EDI_ORDER.tag5_cd = CODE5.code_no";
                sSQL += "                                     and CODE5.code_kbn = 'S'";
                sSQL += "                                     and CODE5.code_sub_kbn = 'B'";
                sSQL += "       left outer join m_SNY_CODE as CODE6 on EDI_ORDER.tag6_cd = CODE6.code_no";
                sSQL += "                                     and CODE6.code_kbn = 'S'";
                sSQL += "                                     and CODE6.code_sub_kbn = 'B'";

                if (!string.IsNullOrEmpty(txtRecieveCnt.Text))
                {
                    chkWhere = true;
                    sRecieveCnt = "EDI_ORDER.received_number = @received_number";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sRecieveCnt;
                }

                if (!string.IsNullOrEmpty(txtRecieveYmd.Text))
                {
                    chkWhere = true;
                    sRecieveYmd = "EDI_ORDER.received_ymd = @received_ymd";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sRecieveYmd;
                }

                if (!string.IsNullOrEmpty(txtShippingYmd.Text))
                {
                    chkWhere = true;
                    sShippingYmd = "EDI_STATUS.shipping_ymd = @shipping_ymd";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sShippingYmd;
                }

                if (!string.IsNullOrEmpty(txtInstNo.Text))
                {
                    chkWhere = true;
                    sInstNo = "EDI_STATUS.inst_no like(@inst_no)";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sInstNo;
                }

                if (!string.IsNullOrEmpty(txtOrderNo.Text))
                {
                    chkWhere = true;
                    sOrderNo = "EDI_STATUS.order_no like(@order_no)";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sOrderNo;
                }

                if (rdoNuno.Checked || rdoSagefuda.Checked)
                {
                    chkWhere = true;
                    sOrderKbn = "EDI_ORDER.order_kbn like(@order_kbn)";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sOrderKbn;
                }

                if (!string.IsNullOrEmpty(txtBrName.Text))
                {
                    chkWhere = true;
                    sBrName = "EDI_ORDER.br_name like (@br_name)";
                    if (string.IsNullOrEmpty(sWhere))
                    {
                        sWhere = "WHERE ";
                    }
                    else
                    {
                        sWhere += "  AND ";
                    }
                    sWhere = sWhere + sBrName;
                }

                if (!chkWhere)
                {
                    MessageBox.Show("検索条件を指定してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!string.IsNullOrEmpty(sWhere))
                {
                    sSQL += sWhere;
                }

                sSQL += " ORDER BY EDI_ORDER.br_name, EDI_STATUS.shipping_ymd, EDI_STATUS.inst_no, EDI_STATUS.order_no";

                sqlSel.CommandText = sSQL;

                if (!string.IsNullOrEmpty(txtRecieveCnt.Text))
                {
                    sqlSel.Parameters.Add(new SqlParameter("@received_number", txtRecieveCnt.Text));
                }

                if (!string.IsNullOrEmpty(txtRecieveYmd.Text))
                {
                    sqlSel.Parameters.Add(new SqlParameter("@received_ymd", txtRecieveYmd.Text.Replace("/", "")));
                }

                if (!string.IsNullOrEmpty(txtShippingYmd.Text))
                {
                    sqlSel.Parameters.Add(new SqlParameter("@shipping_ymd", txtShippingYmd.Text.Replace("/", "")));
                }

                if (!string.IsNullOrEmpty(txtInstNo.Text))
                {
                    sqlSel.Parameters.Add(new SqlParameter("@inst_no", '%' + txtInstNo.Text + "%"));
                }

                if (!string.IsNullOrEmpty(txtOrderNo.Text))
                {
                    sqlSel.Parameters.Add(new SqlParameter("@order_no", '%' + txtOrderNo.Text + "%"));
                }

                if (rdoNuno.Checked)
                {
                    sqlSel.Parameters.Add(new SqlParameter("@order_kbn", '%' + clConst.cdNuno + "%"));
                }

                if (rdoSagefuda.Checked)
                {
                    sqlSel.Parameters.Add(new SqlParameter("@order_kbn", '%' + clConst.cdSagefuda + "%"));
                }

                if (!string.IsNullOrEmpty(txtBrName.Text))
                {
                    sqlSel.Parameters.Add(new SqlParameter("@br_name", '%' + txtBrName.Text + "%"));
                }

                //SQLの実行結果を取得
                SqlDataReader reader = sqlSel.ExecuteReader();

                if (!reader.HasRows)
                {
                    MessageBox.Show("検索条件に該当するデータがありません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    int counter = 0;
                    while (reader.Read())
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dgvEDIOrderList);
                        dgvEDIOrderList.Rows.Add(row);

                        dgvEDIOrderList["chkdetail", counter].Value = false;
                        dgvEDIOrderList["inst_no", counter].Value = "";
                        dgvEDIOrderList["order_no", counter].Value = "";
                        dgvEDIOrderList["order_kbn", counter].Value = "";
                        dgvEDIOrderList["br_name", counter].Value = "";
                        dgvEDIOrderList["shipping_ymd", counter].Value = "";
                        dgvEDIOrderList["received_ymd", counter].Value = "";
                        dgvEDIOrderList["received_number", counter].Value = "";
                        dgvEDIOrderList["tag1_cd", counter].Value = "";
                        dgvEDIOrderList["tag2_cd", counter].Value = "";
                        dgvEDIOrderList["tag3_cd", counter].Value = "";
                        dgvEDIOrderList["tag4_cd", counter].Value = "";
                        dgvEDIOrderList["tag5_cd", counter].Value = "";
                        dgvEDIOrderList["tag6_cd", counter].Value = "";
                        dgvEDIOrderList["tag1_cnt", counter].Value = "";
                        dgvEDIOrderList["tag2_cnt", counter].Value = "";
                        dgvEDIOrderList["tag3_cnt", counter].Value = "";
                        dgvEDIOrderList["tag4_cnt", counter].Value = "";
                        dgvEDIOrderList["tag5_cnt", counter].Value = "";
                        dgvEDIOrderList["tag6_cnt", counter].Value = "";
                        dgvEDIOrderList["identity_flg1", counter].Value = "";
                        dgvEDIOrderList["identity_flg2", counter].Value = "";
                        dgvEDIOrderList["identity_flg3", counter].Value = "";
                        dgvEDIOrderList["identity_flg4", counter].Value = "";
                        dgvEDIOrderList["identity_flg5", counter].Value = "";
                        dgvEDIOrderList["identity_flg6", counter].Value = "";

                        dgvEDIOrderList["inst_no", counter].Value = reader["inst_no"].ToString();
                        dgvEDIOrderList["order_no", counter].Value = reader["order_no"].ToString();
                        dgvEDIOrderList["order_kbn", counter].Value = reader["order_kbn"].ToString(); ;
                        dgvEDIOrderList["br_name", counter].Value = reader["br_name"].ToString();
                        dgvEDIOrderList["shipping_ymd", counter].Value = formatDate(reader["shipping_ymd"].ToString());
                        dgvEDIOrderList["received_ymd", counter].Value = formatDate(reader["received_ymd"].ToString());
                        dgvEDIOrderList["received_number", counter].Value = reader["received_number"].ToString();
                        dgvEDIOrderList["tag1_cd", counter].Value = reader["tag1_cd"].ToString();
                        dgvEDIOrderList["tag2_cd", counter].Value = reader["tag2_cd"].ToString();
                        dgvEDIOrderList["tag3_cd", counter].Value = reader["tag3_cd"].ToString();
                        dgvEDIOrderList["tag4_cd", counter].Value = reader["tag4_cd"].ToString();
                        dgvEDIOrderList["tag5_cd", counter].Value = reader["tag5_cd"].ToString();
                        dgvEDIOrderList["tag6_cd", counter].Value = reader["tag6_cd"].ToString();
                        dgvEDIOrderList["tag1_cnt", counter].Value = reader["tag1_cnt"].ToString();
                        dgvEDIOrderList["tag2_cnt", counter].Value = reader["tag2_cnt"].ToString();
                        dgvEDIOrderList["tag3_cnt", counter].Value = reader["tag3_cnt"].ToString();
                        dgvEDIOrderList["tag4_cnt", counter].Value = reader["tag4_cnt"].ToString();
                        dgvEDIOrderList["tag5_cnt", counter].Value = reader["tag5_cnt"].ToString();
                        dgvEDIOrderList["tag6_cnt", counter].Value = reader["tag6_cnt"].ToString();
                        dgvEDIOrderList["identity_flg1", counter].Value = reader["identity_flg1"].ToString();
                        dgvEDIOrderList["identity_flg2", counter].Value = reader["identity_flg2"].ToString();
                        dgvEDIOrderList["identity_flg3", counter].Value = reader["identity_flg3"].ToString();
                        dgvEDIOrderList["identity_flg4", counter].Value = reader["identity_flg4"].ToString();
                        dgvEDIOrderList["identity_flg5", counter].Value = reader["identity_flg5"].ToString();
                        dgvEDIOrderList["identity_flg6", counter].Value = reader["identity_flg6"].ToString();

                        counter++;
                    }
                    lblResult.Text = counter.ToString() + " 件";
                    sDisplay_flg = clConst.cOn;
                }

                txtRecieveCnt.Enabled = false;
                txtRecieveYmd.Enabled = false;
                txtShippingYmd.Enabled = false;
                txtOrderNo.Enabled = false;
                txtInstNo.Enabled = false;
                rdoAll.Enabled = false;
                rdoNuno.Enabled = false;
                rdoSagefuda.Enabled = false;
                txtBrName.Enabled = false;
            }
            catch (Exception ex)
            {
                // エラーメッセージを出力
                clError.outErrorMsg("9", ex.Message);
            }
            finally
            {
                // DB接続を閉じる
                sqlCon.Close();
                sqlCon.Dispose();
            }
        }
        #endregion

        #region CSV出力ボタン押下時
        /// <summary>
        /// CSV出力ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCSV_Click(object sender, EventArgs e)
        {
            int targetCnt = 0;
            string folderName = string.Empty;

            if (sDisplay_flg.Equals(clConst.cOff))
            {
                MessageBox.Show("出力対象データが表示されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            for (int i = 0; i < dgvEDIOrderList.RowCount; i++)
            {
                DataGridViewCheckBoxCell targetChk = (DataGridViewCheckBoxCell)dgvEDIOrderList["chkdetail", i];
                if (targetChk.Value.ToString().Equals(clConst.cOn))
                {
                    targetCnt++;
                }
            }

            if (targetCnt == 0)
            {
                MessageBox.Show("出力対象データを選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "出力先フォルダを指定してください。";
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.SelectedPath = @"Z:\" + clVariable.sLoginName;
            fbd.ShowNewFolderButton = true;

            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                folderName = fbd.SelectedPath;

                // 指定フォルダ配下に受信日時でフォルダを作成する
                folderName = folderName + clConst.cListTemplate_EDIRainbowCSV
                           + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" +clVariable.sLoginID;
                if (!Directory.Exists(folderName))
                {
                    // フォルダ作成
                    try
                    {
                        Directory.CreateDirectory(folderName);
                    }
                    catch (SystemException)
                    {
                        MessageBox.Show("フォルダの作成に失敗しました。" + clConst.cBR + clConst.cBR + "別のフォルダを指定してください。",
                            "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                return;
            }

            if (outCsvData(folderName)) Form_Initialize();
        }
        #endregion

        #region クリアボタン押下時
        /// <summary>
        /// クリアボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClear_Click(object sender, EventArgs e)
        {
            txtRecieveCnt.Enabled = true;
            txtRecieveYmd.Enabled = true;
            txtShippingYmd.Enabled = true;
            txtInstNo.Enabled = true;
            txtInstNo.Text = string.Empty;
            txtOrderNo.Enabled = true;
            txtOrderNo.Text = string.Empty;
            txtRecieveCnt.Text = string.Empty;
            txtRecieveYmd.Text = string.Empty;
            txtShippingYmd.Text = string.Empty;
            txtBrName.Text = string.Empty;
            dgvEDIOrderList_Initialize();
            lblResult.Text = string.Empty;
            sDisplay_flg = clConst.cOff;

            rdoAll.Enabled = true;
            rdoNuno.Enabled = true;
            rdoSagefuda.Enabled = true;
            rdoAll.Checked = true;
            txtBrName.Enabled = true;
        }
        #endregion

        #region 終了ボタン押下時
        /// <summary>
        /// 終了ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        #region「×」ボタン押下処理
        /// <summary>
        /// 「×」ボタン押下処理
        /// </summary>
        /// <param name="e"></param>
        /// <param name="sender"></param>
        private void frmSNYEDIRainbowCSV_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Owner == null)
            {
                // メニューより起動された場合
                if (clVariable.sOpenStatus == "")
                {
                    // メインメニュー表示
                    frmMainMenu cForm = new frmMainMenu();
                    cForm.Show();
                }
                // メニュー以外より起動された場合
                else
                {
                    clVariable.sOpenStatus = "";
                }
            }
            else
            {
                //order.btnClear_Click(sender, e);
                order.Show();
            }
        }
        #endregion

        #region 画面上でのキーアクション
        /// <summary>
        /// 画面上でのキーアクション
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmSNYEDIOrderSearch_KeyDown(object sender, KeyEventArgs e)
        {
            // キーアクションによる処理
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    btnClose_Click(sender, e);
                    break;
            }
        }
        #endregion

        #region テキストボックス 日付制御
        /// <summary>
        /// テキストボックス 日付制御
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBoxDate_Leave(object sender, EventArgs e)
        {
            string sObjText = ((TextBox)sender).Text;

            string returnText = formatDate(sObjText);
            switch (returnText)
            {
                case "":
                    break;
                case "9":
                    MessageBox.Show("正しい日付の形で入力してください" + clConst.cBR
                        + "入力形式：YYYY/MM/DD", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ((TextBox)sender).Focus();
                    break;
                default:
                    ((TextBox)sender).Text = returnText;
                    break;
            }
        }
        #endregion

        #region 日付項目フォーマット
        /// <summary>
        /// 日付項目フォーマット
        /// </summary>
        /// <param name="sObjText"></param>
        /// <returns>string</returns>
        private string formatDate(string sObjText)
        {
            if (sObjText == "")
            {
                return "";
            }

            if (clModule.IsNumeric(sObjText))
            {
                switch (sObjText.Length)
                {
                    case 4:// mm/dd
                        sObjText = sObjText.Insert(2, "/");
                        break;
                    case 6:// yy/mm/dd
                        sObjText = sObjText.Insert(2, "/");
                        sObjText = sObjText.Insert(5, "/");
                        break;

                    case 8:// yyyy/mm/dd       
                        sObjText = sObjText.Insert(4, "/");
                        sObjText = sObjText.Insert(7, "/");
                        break;
                }
            }

            if (DateTime.TryParse(sObjText, out DateTime dt) == true)
            {
                return clModule.DateYMD(sObjText, "yyyy/MM/dd");
            }
            else
            {
                return "9";
            }
        }
        #endregion

        #region CSV出力処理
        /// <summary>
        /// CSV出力処理
        /// </summary>
        /// <param name="folderName"></param>
        private bool outCsvData(string folderName)
        {
            bool bRetCd = true;
            string inst_no = string.Empty;
            string order_no = string.Empty;
            string order_kbn = string.Empty;
            string errorMessage = string.Empty;
            var orderKbnTable = new List<string>();
            var compotisionTable = new List<string>();

            try
            {
                //DB 接続
                clConnect.getConnect();
                sqlCon = new System.Data.SqlClient.SqlConnection(clConnect.sCn);
                sqlCon.Open();
                sqlSelCsv = sqlCon.CreateCommand();

                StringBuilder sb = new StringBuilder();

                for (int i = 0; i < dgvEDIOrderList.RowCount; i++)
                {
                    DataGridViewCheckBoxCell targetChk = (DataGridViewCheckBoxCell)dgvEDIOrderList["chkdetail", i];

                    if (targetChk.Value.ToString().Equals(clConst.cOn))
                    {
                        string[] tagCd
                            = new string[] { dgvEDIOrderList["tag1_cd", i].Value.ToString(),
                                         dgvEDIOrderList["tag2_cd", i].Value.ToString(),
                                         dgvEDIOrderList["tag3_cd", i].Value.ToString(),
                                         dgvEDIOrderList["tag4_cd", i].Value.ToString(),
                                         dgvEDIOrderList["tag5_cd", i].Value.ToString(),
                                         dgvEDIOrderList["tag6_cd", i].Value.ToString() };

                        string[] tagCnt
                            = new string[] { dgvEDIOrderList["tag1_cnt", i].Value.ToString(),
                                         dgvEDIOrderList["tag2_cnt", i].Value.ToString(),
                                         dgvEDIOrderList["tag3_cnt", i].Value.ToString(),
                                         dgvEDIOrderList["tag4_cnt", i].Value.ToString(),
                                         dgvEDIOrderList["tag5_cnt", i].Value.ToString(),
                                         dgvEDIOrderList["tag6_cnt", i].Value.ToString() };

                        string[] tagIdentity
                            = new string[] { dgvEDIOrderList["identity_flg1", i].Value.ToString(),
                                         dgvEDIOrderList["identity_flg2", i].Value.ToString(),
                                         dgvEDIOrderList["identity_flg3", i].Value.ToString(),
                                         dgvEDIOrderList["identity_flg4", i].Value.ToString(),
                                         dgvEDIOrderList["identity_flg5", i].Value.ToString(),
                                         dgvEDIOrderList["identity_flg6", i].Value.ToString() };

                        inst_no = dgvEDIOrderList["inst_no", i].Value.ToString();
                        order_no = dgvEDIOrderList["order_no", i].Value.ToString();
                        order_kbn = dgvEDIOrderList["order_kbn", i].Value.ToString();

                        if (order_kbn.Contains(clConst.cdSagefuda))
                        {
                            orderKbnTable.Add(clConst.cdSagefuda);
                            compotisionTable.Add("t_SNY_EDI_ORDER_QLTY_S");
                        }

                        if (order_kbn.Contains(clConst.cdNuno))
                        {
                            orderKbnTable.Add(clConst.cdNuno);
                            compotisionTable.Add("t_SNY_EDI_ORDER_QLTY_N");
                        }

                        for (int idx = 0; idx < compotisionTable.Count; idx++)
                        {
                            #region CSV出力対象データ取得SQL
                            sb.Clear();
                            sb.AppendLine("SELECT A.order_no, A.inst_no, order_kbn,");
                            sb.AppendLine("       left(goods_cd, 5) + '-' + substring(goods_cd, 6, 3) as goods_cd,");
                            sb.AppendLine("       substring(goods_cd, 9, 2) as goods_cd_shape,");
                            sb.AppendLine("       season, factory_cd, cloth_name,");
                            sb.AppendLine("       laundry1_item, laundry11_cd, laundry12_cd, laundry13_cd,");
                            sb.AppendLine("       laundry2_item, laundry21_cd, laundry22_cd, laundry23_cd,");
                            sb.AppendLine("       laundry3_item, laundry31_cd, laundry32_cd, laundry33_cd,");
                            sb.AppendLine("       laundry4_item, laundry41_cd, laundry42_cd, laundry43_cd,");
                            sb.AppendLine("       adding1_cd, adding2_cd, adding3_cd, adding4_cd, adding5_cd,");
                            sb.AppendLine("       adding6_cd, adding7_cd, adding8_cd, adding9_cd, adding10_cd,");
                            sb.AppendLine("       adding11_cd, adding12_cd, adding13_cd, adding14_cd, adding15_cd,");
                            sb.AppendLine("       item1,");
                            sb.AppendLine("       disp_item1,");
                            sb.AppendLine("       part1,");
                            sb.AppendLine("       composition1,");
                            sb.AppendLine("       case mixing_ratio1");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio1");
                            sb.AppendLine("       end as mixing_ratio1,");
                            sb.AppendLine("       item2,");
                            sb.AppendLine("       disp_item2,");
                            sb.AppendLine("       part2,");
                            sb.AppendLine("       composition2,");
                            sb.AppendLine("       case mixing_ratio2");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio2");
                            sb.AppendLine("       end as mixing_ratio2,");
                            sb.AppendLine("       item3,");
                            sb.AppendLine("       disp_item3,");
                            sb.AppendLine("       part3,");
                            sb.AppendLine("       composition3,");
                            sb.AppendLine("       case mixing_ratio3");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio3");
                            sb.AppendLine("       end as mixing_ratio3,");
                            sb.AppendLine("       item4,");
                            sb.AppendLine("       disp_item4,");
                            sb.AppendLine("       part4,");
                            sb.AppendLine("       composition4,");
                            sb.AppendLine("       case mixing_ratio4");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio4");
                            sb.AppendLine("       end as mixing_ratio4,");
                            sb.AppendLine("       item5,");
                            sb.AppendLine("       disp_item5,");
                            sb.AppendLine("       part5,");
                            sb.AppendLine("       composition5,");
                            sb.AppendLine("       case mixing_ratio5");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio5");
                            sb.AppendLine("       end as mixing_ratio5,");
                            sb.AppendLine("       item6,");
                            sb.AppendLine("       disp_item6,");
                            sb.AppendLine("       part6,");
                            sb.AppendLine("       composition6,");
                            sb.AppendLine("       case mixing_ratio6");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio6");
                            sb.AppendLine("       end as mixing_ratio6,");
                            sb.AppendLine("       item7,");
                            sb.AppendLine("       disp_item7,");
                            sb.AppendLine("       part7,");
                            sb.AppendLine("       composition7,");
                            sb.AppendLine("       case mixing_ratio7");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio7");
                            sb.AppendLine("       end as mixing_ratio7,");
                            sb.AppendLine("       item8,");
                            sb.AppendLine("       disp_item8,");
                            sb.AppendLine("       part8,");
                            sb.AppendLine("       composition8,");
                            sb.AppendLine("       case mixing_ratio8");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio8");
                            sb.AppendLine("       end as mixing_ratio8,");
                            sb.AppendLine("       item9,");
                            sb.AppendLine("       disp_item9,");
                            sb.AppendLine("       part9,");
                            sb.AppendLine("       composition9,");
                            sb.AppendLine("       case mixing_ratio9");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio9");
                            sb.AppendLine("       end as mixing_ratio9,");
                            sb.AppendLine("       item10,");
                            sb.AppendLine("       disp_item10,");
                            sb.AppendLine("       part10,");
                            sb.AppendLine("       composition10,");
                            sb.AppendLine("       case mixing_ratio10");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio10");
                            sb.AppendLine("       end as mixing_ratio10,");
                            sb.AppendLine("       item11,");
                            sb.AppendLine("       disp_item11,");
                            sb.AppendLine("       part11,");
                            sb.AppendLine("       composition11,");
                            sb.AppendLine("       case mixing_ratio11");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio11");
                            sb.AppendLine("       end as mixing_ratio11,");
                            sb.AppendLine("       item12,");
                            sb.AppendLine("       disp_item12,");
                            sb.AppendLine("       part12,");
                            sb.AppendLine("       composition12,");
                            sb.AppendLine("       case mixing_ratio12");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio12");
                            sb.AppendLine("       end as mixing_ratio12,");
                            sb.AppendLine("       item13,");
                            sb.AppendLine("       disp_item13,");
                            sb.AppendLine("       part13,");
                            sb.AppendLine("       composition13,");
                            sb.AppendLine("       case mixing_ratio13");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio13");
                            sb.AppendLine("       end as mixing_ratio13,");
                            sb.AppendLine("       item14,");
                            sb.AppendLine("       disp_item14,");
                            sb.AppendLine("       part14,");
                            sb.AppendLine("       composition14,");
                            sb.AppendLine("       case mixing_ratio14");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio14");
                            sb.AppendLine("       end as mixing_ratio14,");
                            sb.AppendLine("       item15,");
                            sb.AppendLine("       disp_item15,");
                            sb.AppendLine("       part15,");
                            sb.AppendLine("       composition15,");
                            sb.AppendLine("       case mixing_ratio15");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio15");
                            sb.AppendLine("       end as mixing_ratio15,");
                            sb.AppendLine("       item16,");
                            sb.AppendLine("       disp_item16,");
                            sb.AppendLine("       part16,");
                            sb.AppendLine("       composition16,");
                            sb.AppendLine("       case mixing_ratio16");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio16");
                            sb.AppendLine("       end as mixing_ratio16,");
                            sb.AppendLine("       item17,");
                            sb.AppendLine("       disp_item17,");
                            sb.AppendLine("       part17,");
                            sb.AppendLine("       composition17,");
                            sb.AppendLine("       case mixing_ratio17");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio17");
                            sb.AppendLine("       end as mixing_ratio17,");
                            sb.AppendLine("       item18,");
                            sb.AppendLine("       disp_item18,");
                            sb.AppendLine("       part18,");
                            sb.AppendLine("       composition18,");
                            sb.AppendLine("       case mixing_ratio18");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio18");
                            sb.AppendLine("       end as mixing_ratio18,");
                            sb.AppendLine("       item19,");
                            sb.AppendLine("       disp_item19,");
                            sb.AppendLine("       part19,");
                            sb.AppendLine("       composition19,");
                            sb.AppendLine("       case mixing_ratio19");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio19");
                            sb.AppendLine("       end as mixing_ratio19,");
                            sb.AppendLine("       item20,");
                            sb.AppendLine("       disp_item20,");
                            sb.AppendLine("       part20,");
                            sb.AppendLine("       composition20,");
                            sb.AppendLine("       case mixing_ratio20");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio20");
                            sb.AppendLine("       end as mixing_ratio20,");
                            sb.AppendLine("       item21,");
                            sb.AppendLine("       disp_item21,");
                            sb.AppendLine("       part21,");
                            sb.AppendLine("       composition21,");
                            sb.AppendLine("       case mixing_ratio21");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio21");
                            sb.AppendLine("       end as mixing_ratio21,");
                            sb.AppendLine("       item22,");
                            sb.AppendLine("       disp_item22,");
                            sb.AppendLine("       part22,");
                            sb.AppendLine("       composition22,");
                            sb.AppendLine("       case mixing_ratio22");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio22");
                            sb.AppendLine("       end as mixing_ratio22,");
                            sb.AppendLine("       item23,");
                            sb.AppendLine("       disp_item23,");
                            sb.AppendLine("       part23,");
                            sb.AppendLine("       composition23,");
                            sb.AppendLine("       case mixing_ratio23");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio23");
                            sb.AppendLine("       end as mixing_ratio23,");
                            sb.AppendLine("       item24,");
                            sb.AppendLine("       disp_item24,");
                            sb.AppendLine("       part24,");
                            sb.AppendLine("       composition24,");
                            sb.AppendLine("       case mixing_ratio24");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio24");
                            sb.AppendLine("       end as mixing_ratio24,");
                            sb.AppendLine("       item25,");
                            sb.AppendLine("       disp_item25,");
                            sb.AppendLine("       part25,");
                            sb.AppendLine("       composition25,");
                            sb.AppendLine("       case mixing_ratio25");
                            sb.AppendLine("           when '0' then ''");
                            sb.AppendLine("           else mixing_ratio25");
                            sb.AppendLine("       end as mixing_ratio25,");
                            sb.AppendLine("       country_of_origin1, country_of_origin2, manufac_mk, jacket_price,");
                            sb.AppendLine("       measure_point11, measure_point12, measure_point13, measure_point14, measure_point15,");
                            sb.AppendLine("       measure_point21, measure_point22, measure_point23, measure_point24, measure_point25,");
                            sb.AppendLine("       measure_point31, measure_point32, measure_point33, measure_point34, measure_point35,");
                            sb.AppendLine("       measure_point41, measure_point42, measure_point43, measure_point44, measure_point45,");
                            sb.AppendLine("       measure_point51, measure_point52, measure_point53, measure_point54, measure_point55,");
                            sb.AppendLine("       measure_point61, measure_point62, measure_point63, measure_point64, measure_point65,");
                            sb.AppendLine("       measure_point71, measure_point72, measure_point73, measure_point74, measure_point75,");
                            sb.AppendLine("       measure_point81, measure_point82, measure_point83, measure_point84, measure_point85,");
                            sb.AppendLine("       dimension_value11, dimension_value12, dimension_value13, dimension_value14, dimension_value15,");
                            sb.AppendLine("       dimension_value21, dimension_value22, dimension_value23, dimension_value24, dimension_value25,");
                            sb.AppendLine("       dimension_value31, dimension_value32, dimension_value33, dimension_value34, dimension_value35,");
                            sb.AppendLine("       dimension_value41, dimension_value42, dimension_value43, dimension_value44, dimension_value45,");
                            sb.AppendLine("       dimension_value51, dimension_value52, dimension_value53, dimension_value54, dimension_value55,");
                            sb.AppendLine("       dimension_value61, dimension_value62, dimension_value63, dimension_value64, dimension_value65,");
                            sb.AppendLine("       dimension_value71, dimension_value72, dimension_value73, dimension_value74, dimension_value75,");
                            sb.AppendLine("       dimension_value81, dimension_value82, dimension_value83, dimension_value84, dimension_value85,");
                            sb.AppendLine("       g1_size11, g1_size12, g1_size13, g1_size14, g1_size15, g1_size16, g1_size17, g1_size18,");
                            sb.AppendLine("       g1_prd1_color_no,");
                            sb.AppendLine("       g1_count11, g1_jan11_cd, g1_color11_size,");
                            sb.AppendLine("       g1_count12, g1_jan12_cd, g1_color12_size,");
                            sb.AppendLine("       g1_count13, g1_jan13_cd, g1_color13_size,");
                            sb.AppendLine("       g1_count14, g1_jan14_cd, g1_color14_size,");
                            sb.AppendLine("       g1_count15, g1_jan15_cd, g1_color15_size,");
                            sb.AppendLine("       g1_count16, g1_jan16_cd, g1_color16_size,");
                            sb.AppendLine("       g1_count17, g1_jan17_cd, g1_color17_size,");
                            sb.AppendLine("       g1_count18, g1_jan18_cd, g1_color18_size,");
                            sb.AppendLine("       g2_prd2_color_no,");
                            sb.AppendLine("       g2_count21, g2_jan21_cd, g2_color21_size,");
                            sb.AppendLine("       g2_count22, g2_jan22_cd, g2_color22_size,");
                            sb.AppendLine("       g2_count23, g2_jan23_cd, g2_color23_size,");
                            sb.AppendLine("       g2_count24, g2_jan24_cd, g2_color24_size,");
                            sb.AppendLine("       g2_count25, g2_jan25_cd, g2_color25_size,");
                            sb.AppendLine("       g2_count26, g2_jan26_cd, g2_color26_size,");
                            sb.AppendLine("       g2_count27, g2_jan27_cd, g2_color27_size,");
                            sb.AppendLine("       g2_count28, g2_jan28_cd, g2_color28_size,");
                            sb.AppendLine("       g3_prd3_color_no,");
                            sb.AppendLine("       g3_count31, g3_jan31_cd, g3_color31_size,");
                            sb.AppendLine("       g3_count32, g3_jan32_cd, g3_color32_size,");
                            sb.AppendLine("       g3_count33, g3_jan33_cd, g3_color33_size,");
                            sb.AppendLine("       g3_count34, g3_jan34_cd, g3_color34_size,");
                            sb.AppendLine("       g3_count35, g3_jan35_cd, g3_color35_size,");
                            sb.AppendLine("       g3_count36, g3_jan36_cd, g3_color36_size,");
                            sb.AppendLine("       g3_count37, g3_jan37_cd, g3_color37_size,");
                            sb.AppendLine("       g3_count38, g3_jan38_cd, g3_color38_size,");
                            sb.AppendLine("       g4_prd4_color_no,");
                            sb.AppendLine("       g4_count41, g4_jan41_cd, g4_color41_size,");
                            sb.AppendLine("       g4_count42, g4_jan42_cd, g4_color42_size,");
                            sb.AppendLine("       g4_count43, g4_jan43_cd, g4_color43_size,");
                            sb.AppendLine("       g4_count44, g4_jan44_cd, g4_color44_size,");
                            sb.AppendLine("       g4_count45, g4_jan45_cd, g4_color45_size,");
                            sb.AppendLine("       g4_count46, g4_jan46_cd, g4_color46_size,");
                            sb.AppendLine("       g4_count47, g4_jan47_cd, g4_color47_size,");
                            sb.AppendLine("       g4_count48, g4_jan48_cd, g4_color48_size,");
                            sb.AppendLine("       g5_prd5_color_no,");
                            sb.AppendLine("       g5_count51, g5_jan51_cd, g5_color51_size,");
                            sb.AppendLine("       g5_count52, g5_jan52_cd, g5_color52_size,");
                            sb.AppendLine("       g5_count53, g5_jan53_cd, g5_color53_size,");
                            sb.AppendLine("       g5_count54, g5_jan54_cd, g5_color54_size,");
                            sb.AppendLine("       g5_count55, g5_jan55_cd, g5_color55_size,");
                            sb.AppendLine("       g5_count56, g5_jan56_cd, g5_color56_size,");
                            sb.AppendLine("       g5_count57, g5_jan57_cd, g5_color57_size,");
                            sb.AppendLine("       g5_count58, g5_jan58_cd, g5_color58_size,");
                            sb.AppendLine("       g6_prd6_color_no,");
                            sb.AppendLine("       g6_count61, g6_jan61_cd, g6_color61_size,");
                            sb.AppendLine("       g6_count62, g6_jan62_cd, g6_color62_size,");
                            sb.AppendLine("       g6_count63, g6_jan63_cd, g6_color63_size,");
                            sb.AppendLine("       g6_count64, g6_jan64_cd, g6_color64_size,");
                            sb.AppendLine("       g6_count65, g6_jan65_cd, g6_color65_size,");
                            sb.AppendLine("       g6_count66, g6_jan66_cd, g6_color66_size,");
                            sb.AppendLine("       g6_count67, g6_jan67_cd, g6_color67_size,");
                            sb.AppendLine("       g6_count68, g6_jan68_cd, g6_color68_size,");
                            sb.AppendLine("       tag1_cd, tag2_cd, tag3_cd, tag4_cd, tag5_cd, tag6_cd,");
                            sb.AppendLine("       demerit1_cd, demerit2_cd, demerit3_cd, demerit4_cd, demerit5_cd,");
                            sb.AppendLine("       demerit6_cd, demerit7_cd, demerit8_cd, demerit9_cd, demerit10_cd,");
                            sb.AppendLine("       demerit11_cd, demerit12_cd, demerit13_cd, demerit14_cd, demerit15_cd,");
                            sb.AppendLine("       warning1_cd, warning2_cd, warning3_cd, warning4_cd, warning5_cd,");
                            sb.AppendLine("       warning6_cd, warning7_cd, warning8_cd, warning9_cd, warning10_cd,");
                            sb.AppendLine("       warning11_cd, warning12_cd, warning13_cd, warning14_cd, warning15_cd,");
                            sb.AppendLine("       bottoms_pr1, bottoms_pr2, bottoms_pr3, bottoms_pr4, bottoms_pr5, bottoms_pr6");
                            sb.AppendLine("  FROM t_SNY_EDI_ORDER as A");

                            //品質表示情報は、伝票区分によって参照するテーブルを変更する
                            sb.AppendLine("       INNER JOIN " + compotisionTable[idx].ToString() + " AS B ON A.inst_no = B.inst_no");
                            sb.AppendLine("                                             AND A.order_no = B.order_no");

                            sb.AppendLine(" WHERE A.inst_no = @inst_no AND A.order_no = @order_no");
                            #endregion

                            sqlSelCsv.CommandText = sb.ToString();
                            sqlSelCsv.Parameters.Clear();
                            sqlSelCsv.Parameters.Add(new SqlParameter("@inst_no", inst_no));
                            sqlSelCsv.Parameters.Add(new SqlParameter("@order_no", order_no));

                            //SQLの実行結果を取得
                            using (SqlDataReader reader = sqlSelCsv.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    while (reader.Read())
                                    {
                                        //共通部分の編集
                                        string commonDataHeader = editingCommonPartsHeader(reader);
                                        string commonDataFooter = editingCommonPartsFooter(reader);

                                        //ファイル名(指図書番号 + " " + 伝票NO)
                                        string fileName = string.Empty;

                                        //サイズ色番別データ編集
                                        if (orderKbnTable[idx].ToString().Contains(clConst.cdSagefuda))
                                        {
                                            if (rdoAll.Checked || rdoSagefuda.Checked)
                                            {
                                                //下札の場合
                                                fileName = "S_" + reader["inst_no"].ToString() + "_" + reader["order_no"].ToString();
                                                errorMessage = editingSagefuda(folderName, fileName, reader, commonDataHeader, tagCd, tagCnt, tagIdentity, commonDataFooter);
                                            }
                                        }
                                        else
                                        {
                                            if(rdoAll.Checked || rdoNuno.Checked)
                                            {
                                                //下札以外の場合
                                                fileName = "N_" + reader["inst_no"].ToString() + "_" + reader["order_no"].ToString();
                                                errorMessage = editingNuno(folderName, fileName, reader, commonDataHeader, commonDataFooter);
                                            }
                                        }

                                        if (!errorMessage.Equals(string.Empty))
                                        {
                                            bRetCd = false;
                                            MessageBox.Show(errorMessage, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                            return bRetCd;
                                        }
                                    }
                                }
                            }
                        }
                        orderKbnTable.RemoveRange(0, orderKbnTable.Count);
                        compotisionTable.RemoveRange(0, compotisionTable.Count);
                    }
                }

                MessageBox.Show("ファイルが正常に出力されました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return bRetCd;
            }
            catch (Exception ex)
            {
                clError.outErrorMsg("9", ex.Message);
                return false;
            }
            finally
            {
                // DB接続を閉じる
                sqlCon.Close();
                sqlCon.Dispose();
            }
        }
        #endregion

        #region 共通部分 ヘッダー編集
        /// <summary>
        /// 共通部分 ヘッダー編集
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private string editingCommonPartsHeader(SqlDataReader reader)
        {
            string data = string.Empty;

            data  = reader["goods_cd"].ToString() + "\t";
            data += reader["goods_cd_shape"].ToString() + "\t";
            data += reader["season"].ToString() + "\t";
            data += reader["factory_cd"].ToString() +"\t";
            data += reader["cloth_name"].ToString() +"\t";
            data += reader["laundry1_item"].ToString() +"\t";
            data += reader["laundry11_cd"].ToString() +"\t";
            data += reader["laundry12_cd"].ToString() +"\t";
            data += reader["laundry13_cd"].ToString() +"\t";
            data += reader["laundry2_item"].ToString() +"\t";
            data += reader["laundry21_cd"].ToString() +"\t";
            data += reader["laundry22_cd"].ToString() +"\t";
            data += reader["laundry23_cd"].ToString() +"\t";
            data += reader["laundry3_item"].ToString() +"\t";
            data += reader["laundry31_cd"].ToString() +"\t";
            data += reader["laundry32_cd"].ToString() +"\t";
            data += reader["laundry33_cd"].ToString() +"\t";
            data += reader["laundry4_item"].ToString() +"\t";
            data += reader["laundry41_cd"].ToString() +"\t";
            data += reader["laundry42_cd"].ToString() +"\t";
            data += reader["laundry43_cd"].ToString() +"\t";
            data += reader["adding1_cd"].ToString() + "\t";
            data += reader["adding2_cd"].ToString() + "\t";
            data += reader["adding3_cd"].ToString() + "\t";
            data += reader["adding4_cd"].ToString() + "\t";
            data += reader["adding5_cd"].ToString() + "\t";
            data += reader["adding6_cd"].ToString() + "\t";
            data += reader["adding7_cd"].ToString() + "\t";
            data += reader["adding8_cd"].ToString() + "\t";
            data += reader["adding9_cd"].ToString() + "\t";
            data += reader["adding10_cd"].ToString() + "\t";
            data += reader["adding11_cd"].ToString() + "\t";
            data += reader["adding12_cd"].ToString() + "\t";
            data += reader["adding13_cd"].ToString() + "\t";
            data += reader["adding14_cd"].ToString() + "\t";
            data += reader["adding15_cd"].ToString() + "\t";
            #region 依頼明細票2019-049(編集方法を変更(item2～item25も同様))
            if (reader["item1"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item1"].ToString() + "\t";
            }
            else
            {
                data += reader["item1"].ToString() + ' ' + reader["disp_item1"].ToString() + "\t";
            }
            #endregion
            data += reader["part1"].ToString() + "\t";
            data += reader["composition1"].ToString() + "\t";
            data += reader["mixing_ratio1"].ToString() + "\t";
            if (reader["item2"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item2"].ToString() + "\t";
            }
            else
            {
                data += reader["item2"].ToString() + ' ' + reader["disp_item2"].ToString() + "\t";
            }
            data += reader["part2"].ToString() + "\t";
            data += reader["composition2"].ToString() + "\t";
            data += reader["mixing_ratio2"].ToString() + "\t";
            if (reader["item3"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item3"].ToString() + "\t";
            }
            else
            {
                data += reader["item3"].ToString() + ' ' + reader["disp_item3"].ToString() + "\t";
            }
            data += reader["part3"].ToString() + "\t";
            data += reader["composition3"].ToString() + "\t";
            data += reader["mixing_ratio3"].ToString() + "\t";
            if (reader["item4"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item4"].ToString() + "\t";
            }
            else
            {
                data += reader["item4"].ToString() + ' ' + reader["disp_item4"].ToString() + "\t";
            }
            data += reader["part4"].ToString() + "\t";
            data += reader["composition4"].ToString() + "\t";
            data += reader["mixing_ratio4"].ToString() + "\t";
            if (reader["item5"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item5"].ToString() + "\t";
            }
            else
            {
                data += reader["item5"].ToString() + ' ' + reader["disp_item5"].ToString() + "\t";
            }
            data += reader["part5"].ToString() + "\t";
            data += reader["composition5"].ToString() + "\t";
            data += reader["mixing_ratio5"].ToString() + "\t";
            if (reader["item6"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item6"].ToString() + "\t";
            }
            else
            {
                data += reader["item6"].ToString() + ' ' + reader["disp_item6"].ToString() + "\t";
            }
            data += reader["part6"].ToString() + "\t";
            data += reader["composition6"].ToString() + "\t";
            data += reader["mixing_ratio6"].ToString() + "\t";
            if (reader["item7"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item7"].ToString() + "\t";
            }
            else
            {
                data += reader["item7"].ToString() + ' ' + reader["disp_item7"].ToString() + "\t";
            }
            data += reader["part7"].ToString() + "\t";
            data += reader["composition7"].ToString() + "\t";
            data += reader["mixing_ratio7"].ToString() + "\t";
            if (reader["item8"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item8"].ToString() + "\t";
            }
            else
            {
                data += reader["item8"].ToString() + ' ' + reader["disp_item8"].ToString() + "\t";
            }
            data += reader["part8"].ToString() + "\t";
            data += reader["composition8"].ToString() + "\t";
            data += reader["mixing_ratio8"].ToString() + "\t";
            if (reader["item9"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item9"].ToString() + "\t";
            }
            else
            {
                data += reader["item9"].ToString() + ' ' + reader["disp_item9"].ToString() + "\t";
            }
            data += reader["part9"].ToString() + "\t";
            data += reader["composition9"].ToString() + "\t";
            data += reader["mixing_ratio9"].ToString() + "\t";
            if (reader["item10"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item10"].ToString() + "\t";
            }
            else
            {
                data += reader["item10"].ToString() + ' ' + reader["disp_item10"].ToString() + "\t";
            }
            data += reader["part10"].ToString() + "\t";
            data += reader["composition10"].ToString() + "\t";
            data += reader["mixing_ratio10"].ToString() + "\t";
            if (reader["item11"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item11"].ToString() + "\t";
            }
            else
            {
                data += reader["item11"].ToString() + ' ' + reader["disp_item11"].ToString() + "\t";
            }
            data += reader["part11"].ToString() + "\t";
            data += reader["composition11"].ToString() + "\t";
            data += reader["mixing_ratio11"].ToString() + "\t";
            if (reader["item12"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item12"].ToString() + "\t";
            }
            else
            {
                data += reader["item12"].ToString() + ' ' + reader["disp_item12"].ToString() + "\t";
            }
            data += reader["part12"].ToString() + "\t";
            data += reader["composition12"].ToString() + "\t";
            data += reader["mixing_ratio12"].ToString() + "\t";
            if (reader["item13"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item13"].ToString() + "\t";
            }
            else
            {
                data += reader["item13"].ToString() + ' ' + reader["disp_item13"].ToString() + "\t";
            }
            data += reader["part13"].ToString() + "\t";
            data += reader["composition13"].ToString() + "\t";
            data += reader["mixing_ratio13"].ToString() + "\t";
            if (reader["item14"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item14"].ToString() + "\t";
            }
            else
            {
                data += reader["item14"].ToString() + ' ' + reader["disp_item14"].ToString() + "\t";
            }
            data += reader["part14"].ToString() + "\t";
            data += reader["composition14"].ToString() + "\t";
            data += reader["mixing_ratio14"].ToString() + "\t";
            if (reader["item15"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item15"].ToString() + "\t";
            }
            else
            {
                data += reader["item15"].ToString() + ' ' + reader["disp_item15"].ToString() + "\t";
            }
            data += reader["part15"].ToString() + "\t";
            data += reader["composition15"].ToString() + "\t";
            data += reader["mixing_ratio15"].ToString() + "\t";
            if (reader["item16"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item16"].ToString() + "\t";
            }
            else
            {
                data += reader["item16"].ToString() + ' ' + reader["disp_item16"].ToString() + "\t";
            }
            data += reader["part16"].ToString() + "\t";
            data += reader["composition16"].ToString() + "\t";
            data += reader["mixing_ratio16"].ToString() + "\t";
            if (reader["item17"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item17"].ToString() + "\t";
            }
            else
            {
                data += reader["item17"].ToString() + ' ' + reader["disp_item17"].ToString() + "\t";
            }
            data += reader["part17"].ToString() + "\t";
            data += reader["composition17"].ToString() + "\t";
            data += reader["mixing_ratio17"].ToString() + "\t";
            if (reader["item18"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item18"].ToString() + "\t";
            }
            else
            {
                data += reader["item18"].ToString() + ' ' + reader["disp_item18"].ToString() + "\t";
            }
            data += reader["part18"].ToString() + "\t";
            data += reader["composition18"].ToString() + "\t";
            data += reader["mixing_ratio18"].ToString() + "\t";
            if (reader["item19"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item19"].ToString() + "\t";
            }
            else
            {
                data += reader["item19"].ToString() + ' ' + reader["disp_item19"].ToString() + "\t";
            }
            data += reader["part19"].ToString() + "\t";
            data += reader["composition19"].ToString() + "\t";
            data += reader["mixing_ratio19"].ToString() + "\t";
            if (reader["item20"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item20"].ToString() + "\t";
            }
            else
            {
                data += reader["item20"].ToString() + ' ' + reader["disp_item20"].ToString() + "\t";
            }
            data += reader["part20"].ToString() + "\t";
            data += reader["composition20"].ToString() + "\t";
            data += reader["mixing_ratio20"].ToString() + "\t";
            if (reader["item21"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item21"].ToString() + "\t";
            }
            else
            {
                data += reader["item21"].ToString() + ' ' + reader["disp_item21"].ToString() + "\t";
            }
            data += reader["part21"].ToString() + "\t";
            data += reader["composition21"].ToString() + "\t";
            data += reader["mixing_ratio21"].ToString() + "\t";
            if (reader["item22"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item22"].ToString() + "\t";
            }
            else
            {
                data += reader["item22"].ToString() + ' ' + reader["disp_item22"].ToString() + "\t";
            }
            data += reader["part22"].ToString() + "\t";
            data += reader["composition22"].ToString() + "\t";
            data += reader["mixing_ratio22"].ToString() + "\t";
            if (reader["item23"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item23"].ToString() + "\t";
            }
            else
            {
                data += reader["item23"].ToString() + ' ' + reader["disp_item23"].ToString() + "\t";
            }
            data += reader["part23"].ToString() + "\t";
            data += reader["composition23"].ToString() + "\t";
            data += reader["mixing_ratio23"].ToString() + "\t";
            if (reader["item24"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item24"].ToString() + "\t";
            }
            else
            {
                data += reader["item24"].ToString() + ' ' + reader["disp_item24"].ToString() + "\t";
            }
            data += reader["part24"].ToString() + "\t";
            data += reader["composition24"].ToString() + "\t";
            data += reader["mixing_ratio24"].ToString() + "\t";
            if (reader["item25"].ToString().Equals(string.Empty))
            {
                data += reader["disp_item25"].ToString() + "\t";
            }
            else
            {
                data += reader["item25"].ToString() + ' ' + reader["disp_item25"].ToString() + "\t";
            }
            data += reader["part25"].ToString() + "\t";
            data += reader["composition25"].ToString() + "\t";
            data += reader["mixing_ratio25"].ToString() + "\t";
            data += reader["country_of_origin1"].ToString() + "\t";
            data += reader["country_of_origin2"].ToString() + "\t";
            data += reader["manufac_mk"].ToString() + "\t";
            data += reader["jacket_price"].ToString() + "\t";

            return data;
        }
        #endregion

        #region 共通部分 フッター編集
        /// <summary>
        /// 共通部分 フッター編集
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private string editingCommonPartsFooter(SqlDataReader reader)
        {
            string data = string.Empty;

            data  = reader["tag1_cd"].ToString() + "\t";
            data += reader["tag2_cd"].ToString() + "\t";
            data += reader["tag3_cd"].ToString() + "\t";
            data += reader["tag4_cd"].ToString() + "\t";
            data += reader["tag5_cd"].ToString() + "\t";
            data += reader["tag6_cd"].ToString() + "\t";
            data += reader["demerit1_cd"].ToString() + "\t";
            data += reader["demerit2_cd"].ToString() + "\t";
            data += reader["demerit3_cd"].ToString() + "\t";
            data += reader["demerit4_cd"].ToString() + "\t";
            data += reader["demerit5_cd"].ToString() + "\t";
            data += reader["demerit6_cd"].ToString() + "\t";
            data += reader["demerit7_cd"].ToString() + "\t";
            data += reader["demerit8_cd"].ToString() + "\t";
            data += reader["demerit9_cd"].ToString() + "\t";
            data += reader["demerit10_cd"].ToString() + "\t";
            data += reader["demerit11_cd"].ToString() + "\t";
            data += reader["demerit12_cd"].ToString() + "\t";
            data += reader["demerit13_cd"].ToString() + "\t";
            data += reader["demerit14_cd"].ToString() + "\t";
            data += reader["demerit15_cd"].ToString() + "\t";
            data += reader["warning1_cd"].ToString() + "\t";
            data += reader["warning2_cd"].ToString() + "\t";
            data += reader["warning3_cd"].ToString() + "\t";
            data += reader["warning4_cd"].ToString() + "\t";
            data += reader["warning5_cd"].ToString() + "\t";
            data += reader["warning6_cd"].ToString() + "\t";
            data += reader["warning7_cd"].ToString() + "\t";
            data += reader["warning8_cd"].ToString() + "\t";
            data += reader["warning9_cd"].ToString() + "\t";
            data += reader["warning10_cd"].ToString() + "\t";
            data += reader["warning11_cd"].ToString() + "\t";
            data += reader["warning12_cd"].ToString() + "\t";
            data += reader["warning13_cd"].ToString() + "\t";
            data += reader["warning14_cd"].ToString() + "\t";
            data += reader["warning15_cd"].ToString() + "\t";
            data += reader["bottoms_pr1"].ToString() + "\t";
            data += reader["bottoms_pr2"].ToString() + "\t";
            data += reader["bottoms_pr3"].ToString() + "\t";
            data += reader["bottoms_pr4"].ToString() + "\t";
            data += reader["bottoms_pr5"].ToString() + "\t";
            data += reader["bottoms_pr6"].ToString();

            return data;
        }
        #endregion

        #region サイズ色番別データ編集(下札)
        /// <summary>
        /// サイズ色番別データ編集(下札)
        /// </summary>
        /// <param name="folderName"></param>
        /// <param name="fileName"></param>
        /// <param name="reader"></param>
        /// <param name="header"></param>
        /// <param name="tagCd"></param>
        /// <param name="tagCnt"></param>
        /// <param name="identity"></param>
        /// <param name="footer"></param>
        /// <returns></returns>
        private string editingSagefuda(string folderName, string fileName, SqlDataReader reader, string header, string[] tagCd, string[] tagCnt, string[] identity, string footer)
        {
            int iTagCnt = 0;
            string fIdentity = clConst.cOff;
            string fNormal = clConst.cOff;
            string sSQLIdentity = string.Empty;
            string sDateNow = DateTime.Now.ToString("yyyyMMdd");
            string errorMessage = string.Empty;

            StreamWriter sw = null;

            for (int i = 0; i < 6; i++)
            {
                if (!string.IsNullOrEmpty(tagCd[i].ToString()))
                {
                    //個体識別番号が必要な下札が登録されているかをチェック
                    if (identity[i].ToString().Equals(clConst.cOn))
                    {
                        if (fIdentity.Equals(clConst.cOff))
                        {
                            //個体識別番号が必要な下札で、まだデータ出力していない場合
                            try
                            {
                                //個体識別番号を取得
                                clConnect.getConnect();
                                sqlConIdentity = new System.Data.SqlClient.SqlConnection(clConnect.sCn);
                                sqlConIdentity.Open();
                                sqlIdentity = sqlConIdentity.CreateCommand();
                                sqlIdentity.Transaction = sqlConIdentity.BeginTransaction();

                                // 個体識別番号(最大値 + 1)を取得
                                identityNo = clModule.GetMaxNumberX(clConst.cINDNTNUM, sqlConIdentity, sqlIdentity);

                                string kotaiName = fileName + "_個識" + ".csv";
                                sw = new StreamWriter(folderName + "\\" + kotaiName, false, Encoding.GetEncoding("Shift_JIS"));

                                iTagCnt = int.Parse(tagCnt[i].ToString());

                                //タイトル行の出力
                                createTitle(sw);

                                //明細行の出力
                                createDetail(sw, clConst.cOn, reader, header, footer, clConst.cOn);

                                //個体識別番号を更新
                                sSQLIdentity  = "UPDATE m_COM_NUMBER ";
                                sSQLIdentity += "SET maxnumber = @maxnumber,";
                                sSQLIdentity += "    update_user = @updateUser,";
                                sSQLIdentity += "    update_date = CONVERT(VARCHAR, getdate(), 120) ";
                                sSQLIdentity += "WHERE kbn = @kbn";

                                sqlIdentity.CommandText = sSQLIdentity;

                                sqlIdentity.Parameters.Clear();
                                sqlIdentity.Parameters.Add(new SqlParameter("@maxnumber", identityNo - 1)); //最大値+1が初期値なので-1して更新(イヤだけど)
                                sqlIdentity.Parameters.Add(new SqlParameter("@updateUser", clVariable.sLoginID));
                                sqlIdentity.Parameters.Add(new SqlParameter("@kbn", clConst.cINDNTNUM));

                                sqlIdentity.ExecuteNonQuery();
                                sqlIdentity.Transaction.Commit();
                            }
                            catch (IOException)
                            {
                                sqlIdentity.Transaction.Rollback();
                                errorMessage = "出力するファイルが開かれています。" + clConst.cBR + clConst.cBR + "別のフォルダを指定するか、ファイルを閉じてください。";
                            }
                            catch (SystemException)
                            {
                                sqlIdentity.Transaction.Rollback();
                                errorMessage = "出力するファイルのアクセスが拒否されました。" + clConst.cBR + clConst.cBR + "別のフォルダを指定してください。";
                            }
                            catch (Exception ex)
                            {
                                sqlIdentity.Transaction.Rollback();
                                errorMessage = ex.Message;
                            }
                            finally
                            {
                                sqlConIdentity.Close();
                                sqlConIdentity.Dispose();

                                if (sw != null) sw.Close();
                            }

                            fIdentity = clConst.cOn;
                        }
                    }
                    else
                    {
                        if (fNormal.Equals(clConst.cOff))
                        {
                            //個体識別番号が不要な下札で、まだデータ出力していない場合
                            try
                            {
                                sw = new StreamWriter(folderName + "\\" + fileName + ".csv", false, Encoding.GetEncoding("Shift_JIS"));

                                //タイトル行の出力
                                createTitle(sw);

                                //明細行の出力
                                createDetail(sw, clConst.cOff, reader, header, footer, clConst.cOn);

                                fNormal = clConst.cOn;
                            }
                            catch (IOException)
                            {
                                errorMessage = "出力するファイルが開かれています。" + clConst.cBR + clConst.cBR + "別のフォルダを指定するか、ファイルを閉じてください。";
                            }
                            catch (SystemException)
                            {
                                errorMessage = "出力するファイルのアクセスが拒否されました。" + clConst.cBR + clConst.cBR + "別のフォルダを指定してください。";
                            }
                            catch (Exception ex)
                            {
                                errorMessage = ex.Message;
                            }
                            finally
                            {
                                if (sw != null) sw.Close();
                            }
                        }
                    }
                }
            }

            //仕様追加：個体識別番号を要する下札でもノーマルCSVの出力を行う
            if (fIdentity.Equals(clConst.cOn) && fNormal.Equals(clConst.cOff))
            {
                try
                {
                    sw = new StreamWriter(folderName + "\\" + fileName + ".csv", false, Encoding.GetEncoding("Shift_JIS"));

                    //タイトル行の出力
                    createTitle(sw);

                    //明細行の出力
                    createDetail(sw, clConst.cOff, reader, header, footer, clConst.cOn);
                }
                catch (IOException)
                {
                    errorMessage = "出力するファイルが開かれています。" + clConst.cBR + clConst.cBR + "別のフォルダを指定するか、ファイルを閉じてください。";
                }
                catch (SystemException)
                {
                    errorMessage = "出力するファイルのアクセスが拒否されました。" + clConst.cBR + clConst.cBR + "別のフォルダを指定してください。";
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                }
                finally
                {
                    if (sw != null) sw.Close();
                }
            }

            return errorMessage;
        }
        #endregion

        #region サイズ色番別データ編集(布)
        /// <summary>
        /// サイズ色番別データ編集(布)
        /// </summary>
        /// <param name="folderName"></param>
        /// <param name="fileName"></param>
        /// <param name="reader"></param>
        /// <param name="header"></param>
        /// <param name="footer"></param>
        private string editingNuno(string folderName, string fileName, SqlDataReader reader, string header, string footer)
        {
            StreamWriter sw = null;
            string errorMessage = string.Empty;

            try
            {
                sw = new StreamWriter(folderName + "\\" + fileName + ".csv", false, Encoding.GetEncoding("Shift_JIS"));

                //タイトル行の出力
                createTitle(sw);

                //明細行の出力
                createDetail(sw, clConst.cOff, reader, header, footer, clConst.cOff);
            }
            catch (IOException)
            {
                errorMessage = "出力するファイルが開かれています。" + clConst.cBR + clConst.cBR + "別のフォルダを指定するか、ファイルを閉じてください。";
            }
            catch (SystemException)
            {
                errorMessage = "出力するファイルのアクセスが拒否されました。" + clConst.cBR + clConst.cBR + "別のフォルダを指定してください。";
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            finally
            {
                if (sw != null) sw.Close();
            }

            return errorMessage;
        }
        #endregion

        #region CSVファイルタイトル行出力
        /// <summary>
        /// CSVファイルタイトル行出力
        /// </summary>
        /// <param name="sw"></param>
        private void createTitle(StreamWriter sw)
        {
            string data = string.Empty;

            data  = "商品ｺｰﾄﾞ" + "\t";
            data += "体型" + "\t";
            data += "ｼｰｽﾞﾝ表記" + "\t";
            data += "工場ｺｰﾄﾞ" + "\t";
            data += "ｹｱﾗﾍﾞﾙ素材・幅指定" + "\t";
            data += "洗濯表示(ｱｲﾃﾑ1)" + "\t";
            data += "洗濯表示(表示11)" + "\t";
            data += "洗濯表示(表示12)" + "\t";
            data += "洗濯表示(表示13)" + "\t";
            data += "洗濯表示(ｱｲﾃﾑ2)" + "\t";
            data += "洗濯表示(表示21)" + "\t";
            data += "洗濯表示(表示22)" + "\t";
            data += "洗濯表示(表示23)" + "\t";
            data += "洗濯表示(ｱｲﾃﾑ3)" + "\t";
            data += "洗濯表示(表示31)" + "\t";
            data += "洗濯表示(表示32)" + "\t";
            data += "洗濯表示(表示33)" + "\t";
            data += "洗濯表示(ｱｲﾃﾑ4)" + "\t";
            data += "洗濯表示(表示41)" + "\t";
            data += "洗濯表示(表示42)" + "\t";
            data += "洗濯表示(表示43)" + "\t";
            data += "付記用語1" + "\t";
            data += "付記用語2" + "\t";
            data += "付記用語3" + "\t";
            data += "付記用語4" + "\t";
            data += "付記用語5" + "\t";
            data += "付記用語6" + "\t";
            data += "付記用語7" + "\t";
            data += "付記用語8" + "\t";
            data += "付記用語9" + "\t";
            data += "付記用語10" + "\t";
            data += "付記用語11" + "\t";
            data += "付記用語12" + "\t";
            data += "付記用語13" + "\t";
            data += "付記用語14" + "\t";
            data += "付記用語15" + "\t";
            data += "品質表示(ｱｲﾃﾑ1)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ1)" + "\t";
            data += "品質表示(組成1)" + "\t";
            data += "品質表示(組成混率1)" + "\t";
            data += "品質表示(ｱｲﾃﾑ2)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ2)" + "\t";
            data += "品質表示(組成2)" + "\t";
            data += "品質表示(組成混率2)" + "\t";
            data += "品質表示(ｱｲﾃﾑ3)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ3)" + "\t";
            data += "品質表示(組成3)" + "\t";
            data += "品質表示(組成混率3)" + "\t";
            data += "品質表示(ｱｲﾃﾑ4)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ4)" + "\t";
            data += "品質表示(組成4)" + "\t";
            data += "品質表示(組成混率4)" + "\t";
            data += "品質表示(ｱｲﾃﾑ5)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ5)" + "\t";
            data += "品質表示(組成5)" + "\t";
            data += "品質表示(組成混率5)" + "\t";
            data += "品質表示(ｱｲﾃﾑ6)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ6)" + "\t";
            data += "品質表示(組成6)" + "\t";
            data += "品質表示(組成混率6)" + "\t";
            data += "品質表示(ｱｲﾃﾑ7)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ7)" + "\t";
            data += "品質表示(組成7)" + "\t";
            data += "品質表示(組成混率7)" + "\t";
            data += "品質表示(ｱｲﾃﾑ8)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ8)" + "\t";
            data += "品質表示(組成8)" + "\t";
            data += "品質表示(組成混率8)" + "\t";
            data += "品質表示(ｱｲﾃﾑ9)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ9)" + "\t";
            data += "品質表示(組成9)" + "\t";
            data += "品質表示(組成混率9)" + "\t";
            data += "品質表示(ｱｲﾃﾑ10)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ10)" + "\t";
            data += "品質表示(組成10)" + "\t";
            data += "品質表示(組成混率10)" + "\t";
            data += "品質表示(ｱｲﾃﾑ11)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ11)" + "\t";
            data += "品質表示(組成11)" + "\t";
            data += "品質表示(組成混率11)" + "\t";
            data += "品質表示(ｱｲﾃﾑ12)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ12)" + "\t";
            data += "品質表示(組成12)" + "\t";
            data += "品質表示(組成混率12)" + "\t";
            data += "品質表示(ｱｲﾃﾑ13)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ13)" + "\t";
            data += "品質表示(組成13)" + "\t";
            data += "品質表示(組成混率13)" + "\t";
            data += "品質表示(ｱｲﾃﾑ14)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ14)" + "\t";
            data += "品質表示(組成14)" + "\t";
            data += "品質表示(組成混率14)" + "\t";
            data += "品質表示(ｱｲﾃﾑ15)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ15)" + "\t";
            data += "品質表示(組成15)" + "\t";
            data += "品質表示(組成混率15)" + "\t";
            data += "品質表示(ｱｲﾃﾑ16)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ16)" + "\t";
            data += "品質表示(組成16)" + "\t";
            data += "品質表示(組成混率16)" + "\t";
            data += "品質表示(ｱｲﾃﾑ17)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ17)" + "\t";
            data += "品質表示(組成17)" + "\t";
            data += "品質表示(組成混率17)" + "\t";
            data += "品質表示(ｱｲﾃﾑ18)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ18)" + "\t";
            data += "品質表示(組成18)" + "\t";
            data += "品質表示(組成混率18)" + "\t";
            data += "品質表示(ｱｲﾃﾑ19)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ19)" + "\t";
            data += "品質表示(組成19)" + "\t";
            data += "品質表示(組成混率19)" + "\t";
            data += "品質表示(ｱｲﾃﾑ20)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ20)" + "\t";
            data += "品質表示(組成20)" + "\t";
            data += "品質表示(組成混率20)" + "\t";
            data += "品質表示(ｱｲﾃﾑ21)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ21)" + "\t";
            data += "品質表示(組成21)" + "\t";
            data += "品質表示(組成混率21)" + "\t";
            data += "品質表示(ｱｲﾃﾑ22)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ22)" + "\t";
            data += "品質表示(組成22)" + "\t";
            data += "品質表示(組成混率22)" + "\t";
            data += "品質表示(ｱｲﾃﾑ23)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ23)" + "\t";
            data += "品質表示(組成23)" + "\t";
            data += "品質表示(組成混率23)" + "\t";
            data += "品質表示(ｱｲﾃﾑ24)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ24)" + "\t";
            data += "品質表示(組成24)" + "\t";
            data += "品質表示(組成混率24)" + "\t";
            data += "品質表示(ｱｲﾃﾑ25)" + "\t";
            data += "品質表示(ﾊﾟｰﾂ25)" + "\t";
            data += "品質表示(組成25)" + "\t";
            data += "品質表示(組成混率25)" + "\t";
            data += "原産国1" + "\t";
            data += "原産国2" + "\t";
            data += "製造年月日" + "\t";
            data += "上代" + "\t";
            data += "寸法値1" + "\t";
            data += "寸法値2" + "\t";
            data += "寸法値3" + "\t";
            data += "寸法値4" + "\t";
            data += "寸法値5" + "\t";
            data += "ｻｲｽﾞ" + "\t";
            data += "製品色番" + "\t";
            data += "着数" + "\t";
            data += "JANｺｰﾄﾞ" + "\t";
            data += "印字ｶﾗｰに対するｻｲｽﾞ" + "\t";
            data += "個別識別番号" + "\t";
            data += "下札ｺｰﾄﾞ1" + "\t";
            data += "下札ｺｰﾄﾞ2" + "\t";
            data += "下札ｺｰﾄﾞ3" + "\t";
            data += "下札ｺｰﾄﾞ4" + "\t";
            data += "下札ｺｰﾄﾞ5" + "\t";
            data += "下札ｺｰﾄﾞ6" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号1" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号2" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号3" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号4" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号5" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号6" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号7" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号8" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号9" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号10" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号11" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号12" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号13" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号14" + "\t";
            data += "ﾃﾞﾒﾘｯﾄ表示文章番号15" + "\t";
            data += "注意表示文章番号1" + "\t";
            data += "注意表示文章番号2" + "\t";
            data += "注意表示文章番号3" + "\t";
            data += "注意表示文章番号4" + "\t";
            data += "注意表示文章番号5" + "\t";
            data += "注意表示文章番号6" + "\t";
            data += "注意表示文章番号7" + "\t";
            data += "注意表示文章番号8" + "\t";
            data += "注意表示文章番号9" + "\t";
            data += "注意表示文章番号10" + "\t";
            data += "注意表示文章番号11" + "\t";
            data += "注意表示文章番号12" + "\t";
            data += "注意表示文章番号13" + "\t";
            data += "注意表示文章番号14" + "\t";
            data += "注意表示文章番号15" + "\t";
            data += "組下印字1" + "\t";
            data += "組下印字2" + "\t";
            data += "組下印字3" + "\t";
            data += "組下印字4" + "\t";
            data += "組下印字5" + "\t";
            data += "組下印字6";

            sw.Write(data);
            sw.WriteLine();
        }
        #endregion

        #region CSVファイル明細行出力
        /// <summary>
        /// CSVファイル明細行出力
        /// </summary>
        /// <param name="sw">出力ライター</param>
        /// <param name="identityFlg">個体識別番号フラグ</param>
        /// <param name="reader">EDI SqlDataReader</param>
        /// <param name="header">共通部分ヘッダ</param>
        /// <param name="footer">共通部分フッタ</param>
        /// <param name="sagefudaFlg">下札フラグ</param>
        private void createDetail(StreamWriter sw, string identityFlg, SqlDataReader reader, string header, string footer, string sagefudaFlg)
        {
            char[] separator = new char[] {'|'};

            int iLoopCnt;

            string data = string.Empty;
            string sIdentityNo = string.Empty;
            string wkCount = string.Empty;
            string wkIdentityNo = string.Empty;

            for (int row = 1; row <= rowCnt; row++)
            {
                for (int col = 1; col <= colCnt; col++)
                {
                    string colNo = col.ToString();
                    string rowNo = row.ToString();

                    string measure1 = reader["measure_point" + colNo + "1"].ToString();
                    string measure2 = reader["measure_point" + colNo + "2"].ToString();
                    string measure3 = reader["measure_point" + colNo + "3"].ToString();
                    string measure4 = reader["measure_point" + colNo + "4"].ToString();
                    string measure5 = reader["measure_point" + colNo + "5"].ToString();
                    string dimension1 = reader["dimension_value" + colNo + "1"].ToString();
                    string dimension2 = reader["dimension_value" + colNo + "2"].ToString();
                    string dimension3 = reader["dimension_value" + colNo + "3"].ToString();
                    string dimension4 = reader["dimension_value" + colNo + "4"].ToString();
                    string dimension5 = reader["dimension_value" + colNo + "5"].ToString();
                    string newDimension1 = string.Empty;
                    string newDimension2 = string.Empty;
                    string newDimension3 = string.Empty;
                    string newDimension4 = string.Empty;
                    string newDimension5 = string.Empty;
                    string prdColor = reader["g" + rowNo + "_prd" + rowNo + "_color_no"].ToString();
                    string count = reader["g" + rowNo + "_count" + rowNo + colNo].ToString();
                    string jan = reader["g" + rowNo + "_jan" + rowNo + colNo + "_cd"].ToString();
                    string colorSize = reader["g" + rowNo + "_color" + rowNo + colNo + "_size"].ToString();

                    //寸法値の編集
                    if (sagefudaFlg.Equals(clConst.cOn))
                    {
                        // 下札の場合は "-" を "～" に変換
                        if (!measure1.Equals("呼び名") && !measure1.Equals("呼名"))
                        {
                            newDimension1 = dimension1.Replace('-', '～');
                        }
                        else
                        {
                            newDimension1 = dimension1;
                        }

                        if (!measure2.Equals("呼び名") && !measure2.Equals("呼名"))
                        {
                            newDimension2 = dimension2.Replace('-', '～');
                        }
                        else
                        {
                            newDimension2 = dimension2;
                        }

                        if (!measure3.Equals("呼び名") && !measure3.Equals("呼名"))
                        {
                            newDimension3 = dimension3.Replace('-', '～');
                        }
                        else
                        {
                            newDimension3 = dimension3;
                        }

                        if (!measure4.Equals("呼び名") && !measure4.Equals("呼名"))
                        {
                            newDimension4 = dimension4.Replace('-', '～');
                        }
                        else
                        {
                            newDimension4 = dimension4;
                        }

                        if (!measure5.Equals("呼び名") && !measure5.Equals("呼名"))
                        {
                            newDimension5 = dimension5.Replace('-', '～');
                        }
                        else
                        {
                            newDimension5 = dimension5;
                        }
                    }
                    else
                    {
                        // 布の場合は変換なし
                        newDimension1 = dimension1;
                        newDimension2 = dimension2;
                        newDimension3 = dimension3;
                        newDimension4 = dimension4;
                        newDimension5 = dimension5;
                    }

                    #region 依頼明細票2019-051(データ出力条件を修正)
                    //着数が登録されている場合にデータ出力
                    if (!count.Equals("0"))
                    {
                        if (identityFlg.Equals("1"))
                        {
                            //個体識別番号が必要な伝票は、色番単位の枚数分出力(着数=1)する
                            if (!int.TryParse(count, out iLoopCnt)) iLoopCnt = 0;
                            wkCount = "1";
                        }
                        else
                        {
                            //個体識別番号が不要な伝票は、着数=色番単位の枚数で出力する
                            iLoopCnt = 1;
                            wkCount = count;
                        }

                        for (int i = 1; i <= iLoopCnt; i++)
                        {
                            data = header;
                            data += newDimension1 + "\t";                           //寸法値1
                            data += newDimension2 + "\t";                           //寸法値2
                            data += newDimension3 + "\t";                           //寸法値3
                            data += newDimension4 + "\t";                           //寸法値4
                            data += newDimension5 + "\t";                           //寸法値5
                            data += reader["g1_size1" + colNo].ToString() + "\t";   //サイズ
                            data += prdColor + "\t";                                //製品色番
                            data += wkCount + "\t";                                 //着数
                            data += jan + "\t";                                     //JANコード
                            data += colorSize + "\t";                               //印字カラーに対するサイズ

                            //個体識別番号を編集
                            if (identityFlg.Equals("1"))
                            {
                                wkIdentityNo = identityNo.ToString("0|0|0|0|0|0|0|0|0|0");
                                string[] wkSplited = wkIdentityNo.Split(separator);

                                string[] splitted = new string[5];
                                splitted[0] = wkSplited[8].ToString() + wkSplited[9].ToString();    //分類01
                                splitted[1] = wkSplited[6].ToString() + wkSplited[7].ToString();    //分類02
                                splitted[2] = wkSplited[4].ToString() + wkSplited[5].ToString();    //分類03
                                splitted[3] = wkSplited[2].ToString() + wkSplited[3].ToString();    //分類04
                                splitted[4] = wkSplited[0].ToString() + wkSplited[1].ToString();    //分類05

                                wkIdentityNo = "34" + splitted[4] + splitted[0] + splitted[2] + splitted[3] + splitted[1];
                                identityNo++;
                            }

                            data += wkIdentityNo + "\t";                            //個体識別番号
                            data += footer;
                            sw.Write(data);
                            sw.WriteLine();
                        }
                    }
                    #endregion
                }
            }
        }
        #endregion
    }
}
