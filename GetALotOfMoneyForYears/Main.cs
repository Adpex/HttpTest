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
            if (txt_NowTime.Text.Equals("15:30:00"))
            {
                GetDatas();
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
                
                        DataRow dr = dt.NewRow();
                        dr["债券ID"] = array[i]["cell"]["bond_id"];
                        dr["债券名字"] = array[i]["cell"]["bond_nm"];
                        dr["现价"] = array[i]["cell"]["price"];
                        dr["涨跌幅"] = array[i]["cell"]["increase_rt"];
                        dr["正股ID"] = array[i]["cell"]["pre_bond_id"];
                        dr["正股名称"] = array[i]["cell"]["stock_nm"];
                        dr["正股价"] = array[i]["cell"]["sprice"];
                        dr["正股成交额"] = array[i]["cell"]["svolume"];
                        dr["正股涨跌幅"] = array[i]["cell"]["sincrease_rt"];
                        dr["PB"] = array[i]["cell"]["pb"];
                        dr["转股价"] = array[i]["cell"]["convert_price"];
                        dr["转股价值"] = array[i]["cell"]["convert_value"];
                        dr["转股期提示"] = array[i]["cell"]["convert_cd_tip"];
                        dr["溢价率"] = array[i]["cell"]["premium_rt"];
                        dr["评级"] = array[i]["cell"]["rating_cd"];
                        dr["回售触发价"] = array[i]["cell"]["put_convert_price"];
                        dr["回售起始"] = array[i]["cell"]["next_put_dt"];
                        dr["强赎触发价"] = array[i]["cell"]["force_redeem_price"];
                        dr["转债占比"] = array[i]["cell"]["convert_amt_ratio"];
                        dr["转债规模"] = array[i]["cell"]["orig_iss_amt"];
                        dr["到期时间"] = array[i]["cell"]["short_maturity_dt"];
                        dr["年限"] = array[i]["cell"]["year_left"];
                        dr["到期税前收益"] = array[i]["cell"]["ytm_rt"];
                        dr["到期税后收益"] = array[i]["cell"]["ytm_rt_tax"];
                        dr["成交额"] = array[i]["cell"]["volume"];
                        dt.Rows.Add(dr);
                   
            }

            //建立Excel对象
            Excel.Application excel = new Excel.Application();
            excel.Application.Workbooks.Add(true);
            excel.Visible = true;

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
                    if (dt.Rows[i][j].GetType() == typeof(string))
                    {
                        excel.Cells[i + 2, j + 1] = "'" + dt.Rows[i][j].ToString();
                    }
                    else
                    {
                        excel.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
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
            dcc.Add("债券ID", typeof(String));//债券ID
            dcc.Add("债券名字", typeof(String));//债券名字
            dcc.Add("现价", typeof(String));//现价
            dcc.Add("涨跌幅", typeof(String));//涨跌幅
            dcc.Add("正股ID", typeof(String));//正股ID
            dcc.Add("正股名称", typeof(String));//正股名称
            dcc.Add("正股价", typeof(String));//正股价
            dcc.Add("正股成交额", typeof(String));//正股成交额
            dcc.Add("正股涨跌幅", typeof(String));//正股涨跌幅
            dcc.Add("PB", typeof(String));//PB
            dcc.Add("转股价", typeof(String));//转股价
            dcc.Add("转股价值", typeof(String));//转股价值
            dcc.Add("转股期提示", typeof(String));//转股期提示
            dcc.Add("溢价率", typeof(String));//溢价率
            dcc.Add("评级", typeof(String));//评级
            dcc.Add("回售触发价", typeof(String));//回售触发价
            dcc.Add("回售起始", typeof(String));//回售起始
            dcc.Add("强赎触发价", typeof(String));//强赎触发价
            dcc.Add("转债占比", typeof(String));//转债占比
            dcc.Add("转债规模", typeof(String));//转债规模
            dcc.Add("到期时间", typeof(String));//到期时间
            dcc.Add("年限", typeof(String));//年限
            dcc.Add("到期税前收益", typeof(String));//到期税前收益
            dcc.Add("到期税后收益", typeof(String));//到期税后收益
            dcc.Add("成交额", typeof(String));//成交额
        }
    }
}
