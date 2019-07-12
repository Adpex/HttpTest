using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Excel;
using System.Threading;

namespace GetALotOfMoneyForYears
{
    public partial class Main : Form
    {
        private System.Data.DataTable dt = new System.Data.DataTable();
        public Main()
        {
            InitializeComponent();
        }

        private void Timer_NowTime_Tick(object sender, EventArgs e)
        {
            txt_NowTime.Text = DateTime.Now.ToString("HH:mm:ss");
        }

        private void Main_Load(object sender, EventArgs e)
        {
            CreateDataTables();
        }

        private void Txt_NowTime_TextChanged(object sender, EventArgs e)
        {
            if (txt_NowTime.Text.Equals("15:35:00"))
            {
                ThreadStart ts = new ThreadStart(GetDatas);
                Thread th = new Thread(ts);
                th.Start();
            }
        }

        private void GetDatas()
        {
            string url = "https://www.jisilu.cn/data/cbnew/cb_list/";
            string result = string.Empty;
            Stream dataStream = null;
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                //WebHeaderCollection wc = new WebHeaderCollection();
                //wc.Add("Mozilla/5.0(Macintosh;IntelMacOSX10_7_0)AppleWebKit/535.11(KHTML,likeGecko)Chrome/17.0.963.56Safari/535.11");
                //request.Headers = wc;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                dataStream = response.GetResponseStream();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            if (dataStream != null)
            {
                try
                {
                    using (StreamReader reader = new StreamReader(dataStream, Encoding.UTF8))
                    {
                        result = reader.ReadToEnd();
                    }
                }
                catch (OutOfMemoryException oe)
                {
                    throw new Exception(oe.Data.ToString());
                }
                catch (IOException ie)
                {
                    throw new Exception(ie.Data.ToString());
                }
            }

            JObject jo = JObject.Parse(result);
            JArray array = (JArray)jo["rows"];

            for (int i = 0; i < array.Count; i++)
            {
                string name = array[i]["cell"]["bond_nm"].ToString().Trim();
                double price = GetDoubles(array[i]["cell"]["price"].ToString().Trim());
                string id = array[i]["cell"]["bond_id"].ToString().Trim();
                if (name.Contains("EB"))
                    continue;
                if (price == 0)
                    continue;

                DataRow dr = dt.NewRow();
                dr["债券代码"] = 
                dr["债券名称"] = name;
                
                dr["现价"] = price;
                dr["涨跌幅（%）"] = GetDoubles(array[i]["cell"]["increase_rt"].ToString().Trim().TrimEnd('%'));
                double premium_rt = GetDoubles(array[i]["cell"]["premium_rt"].ToString().Trim().TrimEnd('%'));
                dr["溢价率（%）"] = premium_rt;
                if(array[i]["cell"]["pre_bond_id"].ToString().Substring(0,2).Equals("sh"))
                    dr["证交所"] = "沪";
                else
                    dr["证交所"] = "深";
                dr["正股名称"] = array[i]["cell"]["stock_nm"];
                dr["正股价"] = GetDoubles(array[i]["cell"]["sprice"].ToString().Trim());
                dr["正股成交额（万）"] = GetDoubles(array[i]["cell"]["svolume"].ToString().Trim());
                dr["正股涨跌幅（%）"] = GetDoubles(array[i]["cell"]["sincrease_rt"].ToString().Trim().TrimEnd('%'));
                dr["PB"] = GetDoubles(array[i]["cell"]["pb"].ToString().Trim());
                dr["转股价"] = GetDoubles(array[i]["cell"]["convert_price"].ToString().Trim());
                dr["转股价值"] = GetDoubles(array[i]["cell"]["convert_value"].ToString().Trim());
                string[] spStr = array[i]["cell"]["convert_cd_tip"].ToString().Split('；');
                dr["转股代码"] = spStr[0];
                dr["转股日期"] = Convert.ToDateTime(spStr[1].Substring(0, 10));
                dr["评级"] = array[i]["cell"]["rating_cd"];
                dr["回售触发价"] = GetDoubles(array[i]["cell"]["put_convert_price"].ToString().Trim().TrimEnd('%'));
                string dates = GetDates(array[i]["cell"]["next_put_dt"].ToString().Trim());

                if (!string.IsNullOrEmpty(dates))
                {
                    dr["回售起始"] = Convert.ToDateTime(dates);
                    dr["距回售期月数"] = ((Convert.ToDateTime(dates).Year - DateTime.Now.Year) * 12) + Convert.ToDateTime(dates).Month - DateTime.Now.Month;
                }
                else
                {
                    dr["回售起始"] = DBNull.Value;
                    dr["距回售期月数"] = "无回售条款";
                }
                dr["强赎触发价"] = GetDoubles(array[i]["cell"]["force_redeem_price"].ToString().Trim());
                dr["转债占比（%）"] = GetDoubles(array[i]["cell"]["convert_amt_ratio"].ToString().Trim().TrimEnd('%'));
                dr["转债规模（亿）"] = GetDoubles(array[i]["cell"]["orig_iss_amt"].ToString().Trim().TrimEnd('%'));
                string shortdt = array[i]["cell"]["short_maturity_dt"].ToString().Trim();
                if (!string.IsNullOrEmpty(shortdt))
                    dr["到期时间"] = Convert.ToDateTime("20" + shortdt);
                dr["年限"] = GetDoubles(array[i]["cell"]["year_left"].ToString().Trim());
                dr["到期税前收益（%）"] = GetDoubles(array[i]["cell"]["ytm_rt"].ToString().Trim().TrimEnd('%'));
                dr["到期税后收益（%）"] = GetDoubles(array[i]["cell"]["ytm_rt_tax"].ToString().Trim().TrimEnd('%'));
                dr["成交额（万）"] = GetDoubles(array[i]["cell"]["volume"].ToString().Trim());
                dr["系统价值"] = price + 2 * premium_rt;
                dr["获取时间"] = DateTime.Now.ToString("yyyy-MM-dd");
                dt.Rows.Add(dr);

            }

            //建立Excel对象
            Excel.Application excel = new Excel.Application();
            excel.Application.Workbooks.Add(true);
            excel.Visible = true;
            

            //excel.SaveWorkspace(System.Windows.Forms.Application.StartupPath + string.Format(@"\Data\{0}.xlsx", DateTime.Now.ToString("yyyyMMdd")));
            //excel.Application.SaveWorkspace(System.Windows.Forms.Application.StartupPath + string.Format(@"\Data\{0}.xlsx", DateTime.Now.ToString("yyyyMMdd")));

            //生成字段名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                excel.Cells[1, i + 1] = dt.Columns[i].Caption;
            }
            
            //填充数据
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    //if (dt.Rows[i][j].GetType() == typeof(string))
                    //{
                    //    excel.Cells[i + 2, j + 1] = "'" + dt.Rows[i][j].ToString();
                    //}
                    //else
                    //{
                    //    excel.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    //}
                    excel.Cells[i + 2, j + 1] = dt.Rows[i][j];
                    if (dt.Columns[j] == dt.Columns["距回售期月数"])
                    {
                        string str = dt.Rows[i][j].ToString().Trim();
                        double month = GetDoubles(dt.Rows[i][j].ToString());
                        if (!str.Equals("无回售条款") && month < 24)
                            excel.Cells[i + 2, j + 1].Interior.Color = Color.Yellow;
                        if (!str.Equals("无回售条款") && month < 12)
                            excel.Cells[i + 2, j + 1].Interior.Color = Color.Orange;
                        if (!str.Equals("无回售条款") && month < 6)
                            excel.Cells[i + 2, j + 1].Interior.Color = Color.Red;
                    }
                }
            }
            excel.Cells.Font.Size = 8;
            //excel.Cells.AutoFit();
        }

        private void Btu_SelfGet_Click(object sender, EventArgs e)
        {
            GetDatas();
        }

        private void CreateDataTables()
        {
            DataColumnCollection dcc = dt.Columns;
            dcc.Add("债券代码", typeof(String));//债券ID
            dcc.Add("证交所", typeof(String));//证交所
            dcc.Add("正股名称", typeof(String));//正股名称
            dcc.Add("债券名称", typeof(String));//债券名字
            dcc.Add("现价", typeof(Double));//现价
            dcc.Add("溢价率（%）", typeof(Double));//溢价率
            dcc.Add("系统价值", typeof(Double));//算法：现价+300*溢价率
            dcc.Add("涨跌幅（%）", typeof(Double));//涨跌幅
            dcc.Add("正股价", typeof(Double));//正股价
            dcc.Add("正股涨跌幅（%）", typeof(Double));//正股涨跌幅
            dcc.Add("成交额（万）", typeof(Double));//成交额
            dcc.Add("正股成交额（万）", typeof(Double));//正股成交额
            dcc.Add("PB", typeof(Double));//PB
            dcc.Add("转股价", typeof(Double));//转股价
            dcc.Add("转股价值", typeof(Double));//转股价值
            dcc.Add("评级", typeof(String));//评级
            dcc.Add("回售触发价", typeof(Double));//回售触发价
            dcc.Add("回售起始", typeof(DateTime));//回售起始
            dcc.Add("距回售期月数", typeof(String));//回售起始
            dcc.Add("强赎触发价", typeof(Double));//强赎触发价
            dcc.Add("转股代码", typeof(String));
            dcc.Add("转股日期", typeof(DateTime));//转股期提示
            dcc.Add("转债占比（%）", typeof(String));//转债占比
            dcc.Add("转债规模（亿）", typeof(String));//转债规模
            dcc.Add("到期时间", typeof(DateTime));//到期时间
            dcc.Add("年限", typeof(Double));//年限
            dcc.Add("到期税前收益（%）", typeof(Double));//到期税前收益
            dcc.Add("到期税后收益（%）", typeof(Double));//到期税后收益
            dcc.Add("获取时间", typeof(String));//到期税后收益

        }

        private double GetDoubles(string str)
        {
            double result = -1;
            double.TryParse(str, out result);
            return result;
        }

        private string GetDates(string str)
        {
            string result = string.Empty;
            DateTime dt = new DateTime();
            DateTime.TryParse(str, out dt);
            if (dt.Year != 1)
                result = dt.ToString("yyyy-MM-dd");

            return result;
        }
    }
}
