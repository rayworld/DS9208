using DevComponents.DotNetBar;
using Ray.Framework.Config;
using Ray.Framework.Encrypt;
using System;
using System.Data;

namespace DS9208
{
    public partial class Form1 : Office2007Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTable dt = new DataTable();
        string sql = "";
        string info = "";

        #region 事件

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            this.styleManager1.ManagerStyle = (eStyle)Enum.Parse(typeof(eStyle), ConfigHelper.ReadValueByKey(ConfigHelper.ConfigurationFile.AppConfig, "FormStyle"));
        }

        /// <summary>
        /// 生成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX2_Click(object sender, EventArgs e)
        {
            string year = textBoxX1.Text;
            string tableIndex = textBoxX2.Text;
            dt = QRCodeBuilder(year, tableIndex, true);
            this.dataGridViewX1.DataSource = dt;
            dataGridViewX1.Columns["二维码"].Width = 240;

        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            //string qrCode = "";
            //for (int j = 0; j < dataGridViewX1.Rows.Count; j++)
            //{
            //    qrCode = dataGridViewX1.Rows[j].Cells[1].Value.ToString();
            //    bcc.CreateQuickMark(pictureBox1, qrCode);
            //    Image img = pictureBox1.Image;
            //    img.Save("d:/temp/" + textBoxX2.Text + (j + 1).ToString().PadLeft(5,'0') + ".jpg");
            //}
            Ray.Framework.Converter.DataTable2Excel.DataToExcel(dt);
        }

        #endregion

        #region 私有过程
        /// <summary>
        /// 
        /// </summary>
        /// <param name="year">两位数年份</param>
        /// <param name="startNo">起始编号</param>
        /// <param name="endNo">结束编号</param>
        private static void QRCodeBuilder(string year, int startNo, int endNo)
        {
            string tableIndex = getTableIndexByQRCode(year.ToString() + startNo.ToString().PadLeft(7, '0'));
            if (startNo > 0 && endNo > 0)
            {
                for (int i = startNo; i < endNo; i++)
                {
                    string QRCode = year.ToString() + i.ToString().PadLeft(7, '0');
                    //SqlHelper.ExecuteSql("INSERT INTO [dbo].[" + tableIndex + "] ([FQRCode]) VALUES ('" + EncryptHelper.Encrypt("77052300", QRCode) + "')", null);
                }
            }
        }

        /// <summary>
        /// 由于数据量太大，系统采用分库处理
        /// 本过程通过二维码明文找到分库后表名的索引
        /// </summary>
        /// <param name="QRCode">二维码明文</param>
        /// <returns></returns>
        private static string getTableIndexByQRCode(string QRCode)
        {
            string retVal = "";
            if (!string.IsNullOrEmpty(QRCode))
            {
                string QRYear = QRCode.Substring(0, 2);
                string QRTableIndex = QRCode.Substring(2, 2);
                retVal = "t_QRCode20" + QRYear + QRTableIndex;
            }
            return retVal;
        }


        /// <summary>
        /// 生成二维码
        /// </summary>
        /// <param name="year">两位数年份</param>
        /// <param name="tableIndex">表编号</param>
        /// <param name="enryptable">是否加密</param>
        private DataTable QRCodeBuilder(string year, string tableIndex, bool enryptable)
        {
            dt = new DataTable();
            if (!string.IsNullOrEmpty(year) && !string.IsNullOrEmpty(tableIndex))
            {
                DataColumn dtc0 = new DataColumn("序号", typeof(string));
                dt.Columns.Add(dtc0);
                DataColumn dtc1 = new DataColumn("二维码", typeof(string));
                dt.Columns.Add(dtc1);

                string QRCode = "";
                ///for (int i = 0; i < 100000; i++)
                for (int i = 0; i < 100000; i++)
                {
                    QRCode = year + tableIndex + i.ToString().PadLeft(5, '0');
                    if (enryptable)
                    {
                        QRCode = EncryptHelper.Encrypt(QRCode);
                    }

                    //添加数据到DataTable
                    DataRow dr = dt.NewRow();
                    dr["序号"] = (i + 1).ToString();
                    dr["二维码"] = QRCode;
                    dt.Rows.Add(dr);
                    //SqlHelper.ExecuteSql("INSERT INTO [dbo].[" + "t_QRCode20" + year + tableIndex + "] ([FQRCode]) VALUES ('" + QRCode + "')", null);
                }

            }
            return dt;
        }

        #endregion
        
    }
}
