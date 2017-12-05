using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.DBUtility;
using System;

namespace DS9208
{
    public partial class Form9 : Office2007Form
    {
        public Form9()
        {
            InitializeComponent();
        }
        string info = "";
        string sql = "";

        /// <summary>
        /// 查询QRCode的计数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            string startDate = dateTimeInput1.Value.ToString().Substring(0, 10);
            string endDate = dateTimeInput2.Value.ToString().Substring(0, 10);
            if (startDate != "0001-01-01" || endDate != "0001-01-01")
            {
                int startCounter = 0;
                int endCounter = 0;

                sql = string.Format("SELECT TOP 1 [fCounter] FROM [dbo].[t_Counter] WHERE [fDate] >= '{0}' ORDER BY [fDate] ASC ", startDate);
                object objStartCounter = SqlHelper.ExecuteScalar(sql);
                startCounter = objStartCounter != null ? int.Parse(objStartCounter.ToString()) : 0;
                if (startCounter == 0) 
                {
                    DesktopAlert.Show(Utils.H2("请输入有效的开始时间！"));
                }
                
                sql = string.Format("SELECT TOP 1 [fCounter] FROM [dbo].[t_Counter] WHERE [fDate] <= '{0}' ORDER BY [fDate] DESC ", endDate);
                object objEndCounter = SqlHelper.ExecuteScalar(sql);
                endCounter = objEndCounter != null ? int.Parse(objEndCounter.ToString()) : 0;
                if (endCounter == 0)
                {
                    DesktopAlert.Show(Utils.H2("请输入有效的结束时间！"));
                }

                int QRCodeCount = endCounter - startCounter;
                info = Utils.H2(string.Format("共查询到 {0} 条记录", QRCodeCount.ToString()));
                DesktopAlert.Show(info);
            }
            else
            {
                DesktopAlert.Show(Utils.H2("请输入有效的开始时间和结束时间！"));
            }
        }
    }
}
