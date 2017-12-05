using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.DBUtility;
using Ray.Framework.Encrypt;
using System;
using System.Data;
using System.Windows.Forms;


namespace DS9208
{
    public partial class Form7 : Office2007Form
    {
        
        public Form7()
        {
            InitializeComponent();
        }
        DataTable dt = new DataTable();
        string sql = "";
        string info = "";
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form7_Load(object sender, EventArgs e)
        {
            comboBoxEx2.SelectedIndex = 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxX1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //dataGridViewX1.Rows.Clear();
                //得到单据编号
                string billType = comboBoxEx2.SelectedIndex == 0 ? "XOUT" : "QOUT";
                string billNo = billType + textBoxX1.Text;
                sql = string.Format("SELECT [产品名称] AS Disp , [FEntryID] AS Val FROM [dbo].[icstock] WHERE [单据编号] = '{0}'", billNo);
                dt = SqlHelper.ExecuteDataTable(sql);
                DataRow dr = dt.NewRow();
                dr[0] = "";
                dr[1] = 0;
                dt.Rows.InsertAt(dr, 0);
                comboBoxEx1.DataSource = dt;
                comboBoxEx1.DisplayMember = "Disp";
                comboBoxEx1.ValueMember = "Val";
                comboBoxEx1.Focus();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBoxX2.Focus();
        }

        private void textBoxX2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string billType = comboBoxEx2.SelectedIndex == 0 ? "XOUT" : "QOUT";
                string billNo = billType + textBoxX1.Text;
                string QRCode = textBoxX2.Text;
                string mingQRCode = EncryptHelper.Decrypt(QRCode);
                string EntryID = comboBoxEx1.SelectedValue.ToString();
                string interID = billNo + comboBoxEx1.SelectedValue.ToString().PadLeft(4, '0');
                string tableName = "t_QRCode" + mingQRCode.Substring(0, 4);
                int retValDetail = 0;
                int retValTotal = 0;
                sql = string.Format("DELETE FROM " + tableName + "  WHERE [FQRCode] = '" + mingQRCode + "' AND [FEntryID] = '" + interID + "'");
                retValDetail = SqlHelper.ExecuteNonQuery(sql);
                sql = string.Format("UPDATE [icstock] SET [FActQty] = [FActQty] - 1 WHERE  [单据编号] = '{0}' AND [FActQty] > 0 AND [FEntryID] = {1}", billNo, EntryID.ToString());
                retValTotal = SqlHelper.ExecuteNonQuery(sql);
                if (retValTotal > 0 && retValDetail > 0)
                {
                    DesktopAlert.Show(Utils.H2("二维码删除成功！"));
                }
                else if (retValDetail < 1)
                {
                    DesktopAlert.Show(Utils.H2("二维码不存在！"));
                }
                else
                {
                    DesktopAlert.Show(Utils.H2("二维码删除失败！"));
                }

                //清空二维码框
                textBoxX2.Text = "";
            }
        }
    }
}
