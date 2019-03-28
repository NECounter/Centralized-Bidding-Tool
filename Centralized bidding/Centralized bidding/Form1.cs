using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Centralized_bidding
{

    public partial class Main : Form
    {
        
        private double[,] SortedPrice_S;//供应价格排序后的数组，（价格,供应量）
        private double[,] SortedPrice_N;//需求价格排序后的数组，（价格,供应量）

        private Participator[] Plants; //声明电厂
        private Participator[] Needers;//声明需求方

        private int PlantsCount; //电厂总数
        private int NeedersCount;//需求方总数

        private string CurrentFile; //当前 打开的文件

        private double ClearingPriceTemp = 0.0;//出清价格 （平均数）
        private double ClearingPrice_S = 0.0;//出清价格（供应方）
        private double BiasAmount = 0.0;//达到需求时，计算多出来的中间量

        private double TotalAmount = 0.0;//需求总量

        bool Caled = false;//该组数据是否处理过
       

        /// <summary>
        /// 构造函数
        /// </summary>
        public Main()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 窗口载入时执行的操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Main_Load(object sender, EventArgs e)
        {
            /***********************初始化数据表格*********************/
            dataGridView1.Rows.Add(1000);
            for (int i = 0; i < 1000; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = (i+1).ToString();
            }
            dataGridView2.Rows.Add(1000);
            for (int i = 0; i < 1000; i++)
            {
                dataGridView2.Rows[i].Cells[0].Value = (i + 1).ToString();
            }
            /***********************初始化数据表格*********************/
        }

        /// <summary>
        /// 计算按钮按下时执行的操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {
            /***********************检查数据是否有缺失*********************/
            for (int i = 0; i < 1000; i++)
            {
                if(dataGridView1.Rows[i].Cells[1].Value != null)
                {
                    if(
                       dataGridView1.Rows[i].Cells[2].Value == null ||
                       dataGridView1.Rows[i].Cells[3].Value == null ||
                       dataGridView1.Rows[i].Cells[4].Value == null ||
                       dataGridView1.Rows[i].Cells[5].Value == null ||
                       dataGridView1.Rows[i].Cells[6].Value == null ||
                       dataGridView1.Rows[i].Cells[7].Value == null
                       )
                    {
                        MessageBox.Show("供应方“" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "”数据缺失！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
                if(
                    dataGridView1.Rows[i].Cells[2].Value != null ||
                    dataGridView1.Rows[i].Cells[3].Value != null ||
                    dataGridView1.Rows[i].Cells[4].Value != null ||
                    dataGridView1.Rows[i].Cells[5].Value != null ||
                    dataGridView1.Rows[i].Cells[6].Value != null ||
                    dataGridView1.Rows[i].Cells[7].Value != null
                    )
                {
                    if(dataGridView1.Rows[i].Cells[1].Value == null)
                    {
                        MessageBox.Show("请补全供应方名称！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
            }
            for (int i = 0; i < 1000; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value != null)
                {
                    if (
                       dataGridView2.Rows[i].Cells[2].Value == null ||
                       dataGridView2.Rows[i].Cells[3].Value == null ||
                       dataGridView2.Rows[i].Cells[4].Value == null ||
                       dataGridView2.Rows[i].Cells[5].Value == null ||
                       dataGridView2.Rows[i].Cells[6].Value == null ||
                       dataGridView2.Rows[i].Cells[7].Value == null 
                       )
                    {
                        MessageBox.Show("需求方“" + dataGridView2.Rows[i].Cells[1].Value.ToString() + "”数据缺失！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
                if (
                    dataGridView2.Rows[i].Cells[2].Value != null ||
                    dataGridView2.Rows[i].Cells[3].Value != null ||
                    dataGridView2.Rows[i].Cells[4].Value != null ||
                    dataGridView2.Rows[i].Cells[5].Value != null ||
                    dataGridView2.Rows[i].Cells[6].Value != null ||
                    dataGridView2.Rows[i].Cells[7].Value != null
                    )
                {
                    if (dataGridView2.Rows[i].Cells[1].Value == null)
                    {
                        MessageBox.Show("请补需求方名称！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
            }
            /***********************检查数据是否有缺失*********************/


            if (ReadData_S() == false) //读取表格1里的数据到缓存 
            {
                return;//读取失败就跳出
            }

            if (ReadData_N() == false)//读取表格2里的数据到缓存 
            {
                return;//读取失败就跳出
            }

            Sort_S(Plants); //把从表格1读取到缓存的数据按照价格从低到高排序
            Sort_N(Needers);//把从表格2读取到缓存的数据按照价格从高到低排序

            double temp = ClearingPrice(SortedPrice_N, SortedPrice_S);//算出清出价格给temp

            if (temp == 0.0)
            {
                 MessageBox.Show("供小于求"); //弹出对话框
            }
            else
            {
                textBox2.Text = temp.ToString();
            }

            //计算生产者，消费者剩余并显示
            textBox5.Text = ProducerSurplus().ToString("0.000");
            textBox6.Text = ConsumerSurplus().ToString("0.000");

            Caled = true;//标记为已计算过
        }
        /// <summary>
        /// 提供供应价格排序，价格从低到高
        /// </summary>
        private void Sort_S(Participator[] Plants)
        {
            /***************读取Plants缓存里的数据到SortedPrice_S二维数组中*************/
            SortedPrice_S = new double[3 * (PlantsCount), 3];

            for(int i = 0; i < PlantsCount; i++)
            {
                SortedPrice_S[3 * i, 0] = Plants[i].PriceLow;
                SortedPrice_S[3 * i + 1, 0] = Plants[i].PriceMid;
                SortedPrice_S[3 * i + 2, 0] = Plants[i].PriceHigh;
                SortedPrice_S[3 * i, 1] = Plants[i].SupplyLow;
                SortedPrice_S[3 * i + 1, 1] = Plants[i].SupplyMid;
                SortedPrice_S[3 * i + 2, 1] = Plants[i].SupplyHigh;
                SortedPrice_S[3 * i, 2] = i;
                SortedPrice_S[3 * i + 1, 2] = i;
                SortedPrice_S[3 * i + 2, 2] = i;
            }
            /***************读取Plants缓存里的数据到SortedPrice_S二维数组中*************/

            /***************冒泡排序*************/
            for (int i = 3 * PlantsCount - 1; i > 0; i--)
            {
                for (int j = 0; j < i; j++)
                {
                    if (SortedPrice_S[j, 0] > SortedPrice_S[j + 1, 0])
                    {
                        double temp;
                        temp = SortedPrice_S[j, 0];
                        SortedPrice_S[j, 0] = SortedPrice_S[j + 1, 0];
                        SortedPrice_S[j + 1, 0] = temp;

                        temp = SortedPrice_S[j, 1];
                        SortedPrice_S[j, 1] = SortedPrice_S[j + 1, 1];
                        SortedPrice_S[j + 1, 1] = temp;

                        temp = SortedPrice_S[j, 2];
                        SortedPrice_S[j, 2] = SortedPrice_S[j + 1, 2];
                        SortedPrice_S[j + 1, 2] = temp;
                    }

                }
            }
            /***************冒泡排序*************/
        }

        /// <summary>
        /// 提供需求价格排序，从高到低
        /// </summary>
        private void Sort_N(Participator[] Needers)
        {
            /***************读取Needers缓存里的数据到SortedPrice_N二维数组中*************/
            SortedPrice_N = new double[3 * (NeedersCount), 2];

            for (int i = 0; i < NeedersCount; i++)
            {
                SortedPrice_N[3 * i, 0] = Needers[i].PriceLow;
                SortedPrice_N[3 * i + 1, 0] = Needers[i].PriceMid;
                SortedPrice_N[3 * i + 2, 0] = Needers[i].PriceHigh;
                SortedPrice_N[3 * i, 1] = Needers[i].SupplyLow;
                SortedPrice_N[3 * i + 1, 1] = Needers[i].SupplyMid;
                SortedPrice_N[3 * i + 2, 1] = Needers[i].SupplyHigh;
            }
            /***************读取Needers缓存里的数据到SortedPrice_N二维数组中*************/

            /***************冒泡排序*************/
            for (int i = 3 * NeedersCount - 1; i > 0; i--)
            {
                for (int j = 0; j < i; j++)
                {
                    if (SortedPrice_N[j, 0] < SortedPrice_N[j + 1, 0])
                    {
                        double temp;
                        temp = SortedPrice_N[j, 0];
                        SortedPrice_N[j, 0] = SortedPrice_N[j + 1, 0];
                        SortedPrice_N[j + 1, 0] = temp;

                        temp = SortedPrice_N[j, 1];
                        SortedPrice_N[j, 1] = SortedPrice_N[j + 1, 1];
                        SortedPrice_N[j + 1, 1] = temp;
                    }

                }
            }
            /***************冒泡排序*************/
        }

        /// <summary>
        /// 读取供应表格里的数据到缓存的实现
        /// </summary>
        private bool ReadData_S()
        { 
            if(dataGridView1.Rows[0].Cells[1].Value == null)
            {
                MessageBox.Show("没有供应方数据！"); //如果发现第一行没有数据，直接返回读取失败
                return false;
            }


            Plants = new Participator[1000];
            int i = 0;//计数器
            do
            {
                Participator pp = new Participator();
                pp.PlantName = dataGridView1.Rows[i].Cells[1].Value.ToString();
                pp.PriceLow = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value);
                pp.SupplyLow = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                pp.PriceMid = Convert.ToDouble(dataGridView1.Rows[i].Cells[4].Value);
                pp.SupplyMid = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                pp.PriceHigh = Convert.ToDouble(dataGridView1.Rows[i].Cells[6].Value);
                pp.SupplyHigh = Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                pp.TotalCost = Convert.ToDouble(dataGridView1.Rows[i].Cells[8].Value);
                Plants[i] = pp;
                i++;
            } while (dataGridView1.Rows[i].Cells[1].Value != null); //有数据就读
            PlantsCount = i;//读完记录一下有多少个厂
            return true;//返回读取成功
        }

        /// <summary>
        /// 读取需求表格里的数据到缓存
        /// </summary>
        private bool ReadData_N()
        {
            if (dataGridView2.Rows[0].Cells[1].Value == null)
            {
                MessageBox.Show("没有需求方数据！");//如果发现第一行没有数据，直接返回读取失败
                return false;
            }
            Needers = new Participator[1000];
            int i = 0;//计数器
            do
            {
                Participator pp = new Participator();
                pp.PlantName = dataGridView2.Rows[i].Cells[1].Value.ToString();
                pp.PriceLow = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                pp.SupplyLow = Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
                pp.PriceMid = Convert.ToDouble(dataGridView2.Rows[i].Cells[4].Value);
                pp.SupplyMid = Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value);
                pp.PriceHigh = Convert.ToDouble(dataGridView2.Rows[i].Cells[6].Value);
                pp.SupplyHigh = Convert.ToDouble(dataGridView2.Rows[i].Cells[7].Value);
                pp.AimPrice = Convert.ToDouble(dataGridView2.Rows[i].Cells[8].Value);
                Needers[i] = pp;
                i++;
            } while (dataGridView2.Rows[i].Cells[1].Value != null);//有数据就读
            NeedersCount = i; //读完记录一下有多少个需求方
            return true;//返回读取成功
        }

       /// <summary>
       /// 算出出清价格的实现
       /// </summary>
       /// <param name="sortedPrice_n"></param>
       /// <param name="sortedPrice_s"></param>
       /// <returns></returns>
        private double ClearingPrice(double[,] sortedPrice_n, double[,] sortedPrice_s)
        {
            /********局部变量*******/
            double clearingPrice_s = 0.0;//供应方出清价格
            double clearingPrice_n = 0.0;//需求方出清价格
            double currentAmount = 0.0;//当前的计算中间总量
            double totalDemand = 0.0;//总需求
            double totalClearingAmount = 0.0;//总出清量
            /********局部变量*******/

            for (int i = 0; i < 3 * NeedersCount; i++)
            {
                totalDemand += sortedPrice_n[i, 1];//算出总需求量
            }
            TotalAmount = totalDemand;//总需求量给TotalAmount供外部使用

            double[] clearingAmount = new double[PlantsCount];//新建数组，存放每个供应方的出清量

            /*
             * 计算出清量逻辑：将每个价格的供应量按照其供应价格从小到大排序之后，从价钱少的开始累加，当加到
             * 累加总量正好大于或者等于需求总量时，这时候所对应的价格就是供应方出清量
             * 因为是买方市场，所以需求方的出清价格就是其最低的需求价格，只要取排序后的数组的最后一个就行了。
             */
            for (int i = 0; i < 3 * PlantsCount; i++)
            {
                currentAmount += sortedPrice_s[i, 1];//累加
                clearingAmount[Convert.ToInt32((sortedPrice_s[i, 2]))] += sortedPrice_s[i, 1];//累加每个厂的出清量

                for (int ii = 0; ii < PlantsCount; ii++)
                {
                    //显示每个厂的出清量
                    dataGridView1.Rows[ii].Cells[9].Value = clearingAmount[ii].ToString();
                }
            
                
                if (currentAmount >= totalDemand)//一旦大于总需求，就执行，最后结束流程
                {
                    clearingPrice_s = sortedPrice_s[i, 0];
                    ClearingPrice_S = sortedPrice_s[i, 0];
                    textBox4.Text = clearingPrice_s.ToString();
                    clearingPrice_n = sortedPrice_n[3 * NeedersCount - 1, 0];
                    textBox1.Text = clearingPrice_n.ToString();
                    ClearingPriceTemp = (clearingPrice_s + clearingPrice_n) / 2;
                    BiasAmount = (currentAmount - totalDemand);
                    clearingAmount[Convert.ToInt32((sortedPrice_s[i, 2]))] -= BiasAmount;
                    for (int ii = 0; ii < PlantsCount; ii++)
                    {
                        dataGridView1.Rows[ii].Cells[9].Value = clearingAmount[ii].ToString();
                        totalClearingAmount += clearingAmount[ii];
                    }
                    textBox3.Text = totalClearingAmount.ToString();
                  
                    return (clearingPrice_s + clearingPrice_n) / 2;
                }
            }
            
            return 0.0;//供小于求      
        }

        /// <summary>
        /// 计算消费者剩余
        /// </summary>
        /// <returns></returns>
        private double ConsumerSurplus()
        {
            double consumerSurplus = 0.0;
            for(int i = 0; i < NeedersCount; i++)
            {
                consumerSurplus += (Needers[i].AimPrice - Needers[i].PriceLow) * Needers[i].SupplyLow +
                                   (Needers[i].AimPrice - Needers[i].PriceMid) * Needers[i].SupplyMid +
                                   (Needers[i].AimPrice - Needers[i].PriceHigh) * Needers[i].SupplyHigh;
            }
            return consumerSurplus;
        }
        /// <summary>
        /// 计算生产者剩余
        /// </summary>
        /// <returns></returns>
        private double ProducerSurplus()
        {
            double producerSurplus = 0.0;
            for (int i = 0; i < PlantsCount; i++)
            {
                if (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) >= Plants[i].SupplyLow)
                {
                    producerSurplus += (Plants[i].PriceLow - Plants[i].TotalCost) * Plants[i].SupplyLow;
                    if (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) >= (Plants[i].SupplyLow + Plants[i].SupplyMid))
                    {
                        producerSurplus += (Plants[i].PriceMid - Plants[i].TotalCost) * Plants[i].SupplyMid;
                        if (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) >= (Plants[i].SupplyLow + Plants[i].SupplyMid + Plants[i].SupplyHigh))
                        {
                            producerSurplus += (Plants[i].PriceHigh - Plants[i].TotalCost) * Plants[i].SupplyHigh;
                        }
                        else
                        {
                            producerSurplus += (Plants[i].PriceHigh - Plants[i].TotalCost) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) - Plants[i].SupplyLow - Plants[i].SupplyMid);
                        }
                    }
                    else
                    {
                        producerSurplus += (Plants[i].PriceMid - Plants[i].TotalCost) * (Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value) - Plants[i].SupplyLow);
                    }
                }
                else
                {
                    producerSurplus += (Plants[i].PriceLow - Plants[i].TotalCost) * Convert.ToDouble(dataGridView1.Rows[i].Cells[9].Value);
                }         
            }
            return producerSurplus;
        }
        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="filePath"></param>
        private void SaveFile(string filePath)
        {
            StreamWriter sw = new StreamWriter(filePath,false,Encoding.UTF8);
            sw.WriteLine("序号,供应方名称,供应低价,低价供应量,供应中价,中价供应量,供应高价,高价供应量,生产总成本");
            int i = 0;
            while (dataGridView1.Rows[i].Cells[1].Value != null)
            {
                sw.WriteLine(dataGridView1.Rows[i].Cells[0].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[1].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[2].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[3].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[4].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[5].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[6].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[7].Value.ToString() + "," +
                             dataGridView1.Rows[i].Cells[8].Value.ToString()
                             );
                i++;
            }

            sw.WriteLine("Split,,,,,,,,");
            sw.WriteLine("序号,需求方名称,需求低价,低价供应量,需求中价,中价供应量,需求高价,高价供应量,辅助服务心理价格,");
            int j = 0;
            while (dataGridView2.Rows[j].Cells[1].Value != null)
            {
                sw.WriteLine(dataGridView2.Rows[j].Cells[0].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[1].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[2].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[3].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[4].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[5].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[6].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[7].Value.ToString() + "," +
                             dataGridView2.Rows[j].Cells[8].Value.ToString() + ","
                             );
                j++;
            }
            sw.WriteLine("END,,,,,,,,");

            sw.Close();

            MessageBox.Show("保存成功！");

        }
        /// <summary>
        /// 读取文件
        /// </summary>
        /// <param name="filePath"></param>
        private void ReadFile(string filePath)
        {
            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            sr.ReadLine();
            string[] data;
            int i = 0;
            do
            {
                data = sr.ReadLine().Split(',');
                
                if(data[0] != "Split")
                {
                    dataGridView1.Rows[i].Cells[1].Value = data[1];
                    dataGridView1.Rows[i].Cells[2].Value = data[2];
                    dataGridView1.Rows[i].Cells[3].Value = data[3];
                    dataGridView1.Rows[i].Cells[4].Value = data[4];
                    dataGridView1.Rows[i].Cells[5].Value = data[5];
                    dataGridView1.Rows[i].Cells[6].Value = data[6];
                    dataGridView1.Rows[i].Cells[7].Value = data[7];
                    dataGridView1.Rows[i].Cells[8].Value = data[8];
                }
                i++;
            } while (data[0] != "Split");
            sr.ReadLine();
            i = 0;
            do
            {
                data = sr.ReadLine().Split(',');
                if (data[0] != "END")
                {
                    dataGridView2.Rows[i].Cells[1].Value = data[1];
                    dataGridView2.Rows[i].Cells[2].Value = data[2];
                    dataGridView2.Rows[i].Cells[3].Value = data[3];
                    dataGridView2.Rows[i].Cells[4].Value = data[4];
                    dataGridView2.Rows[i].Cells[5].Value = data[5];
                    dataGridView2.Rows[i].Cells[6].Value = data[6];
                    dataGridView2.Rows[i].Cells[7].Value = data[7];
                    dataGridView2.Rows[i].Cells[8].Value = data[8];
                }
                i++;
            } while (data[0] != "END");

            sr.Close();
            MessageBox.Show("读取成功！");

        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "数据文件(*.csv)|*.csv;|所有文件(*.*)|*.*";
            ofd.ShowDialog();
            if(ofd.FileName != "")
            {
                CurrentFile = ofd.FileName;
                for (int i = 0; i < 1000; i++)
                {
                    dataGridView1.Rows[i].Cells[1].Value = null;
                    dataGridView1.Rows[i].Cells[2].Value = null;
                    dataGridView1.Rows[i].Cells[3].Value = null;
                    dataGridView1.Rows[i].Cells[4].Value = null;
                    dataGridView1.Rows[i].Cells[5].Value = null;
                    dataGridView1.Rows[i].Cells[6].Value = null;
                    dataGridView1.Rows[i].Cells[7].Value = null;
                    dataGridView1.Rows[i].Cells[8].Value = null;
                
                    dataGridView2.Rows[i].Cells[1].Value = null;
                    dataGridView2.Rows[i].Cells[2].Value = null;
                    dataGridView2.Rows[i].Cells[3].Value = null;
                    dataGridView2.Rows[i].Cells[4].Value = null;
                    dataGridView2.Rows[i].Cells[5].Value = null;
                    dataGridView2.Rows[i].Cells[6].Value = null;
                    dataGridView2.Rows[i].Cells[7].Value = null;
                    dataGridView2.Rows[i].Cells[8].Value = null;


                }

                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                try
                {
                    ReadFile(CurrentFile);
                    Caled = false;
                }
                catch (Exception)
                {
                    MessageBox.Show("文件无效！");
                   
                }
                
            }
            else
            {
                MessageBox.Show("没有选择文件！");
            }


        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 1000; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value != null)
                {
                    if (
                       dataGridView1.Rows[i].Cells[2].Value == null ||
                       dataGridView1.Rows[i].Cells[3].Value == null ||
                       dataGridView1.Rows[i].Cells[4].Value == null ||
                       dataGridView1.Rows[i].Cells[5].Value == null ||
                       dataGridView1.Rows[i].Cells[6].Value == null ||
                       dataGridView1.Rows[i].Cells[7].Value == null
                       )
                    {
                        MessageBox.Show("供应方“" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "”数据缺失！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
                if (
                    dataGridView1.Rows[i].Cells[2].Value != null ||
                    dataGridView1.Rows[i].Cells[3].Value != null ||
                    dataGridView1.Rows[i].Cells[4].Value != null ||
                    dataGridView1.Rows[i].Cells[5].Value != null ||
                    dataGridView1.Rows[i].Cells[6].Value != null ||
                    dataGridView1.Rows[i].Cells[7].Value != null
                    )
                {
                    if (dataGridView1.Rows[i].Cells[1].Value == null)
                    {
                        MessageBox.Show("请补全供应方名称！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
            }
            for (int i = 0; i < 1000; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value != null)
                {
                    if (
                       dataGridView2.Rows[i].Cells[2].Value == null ||
                       dataGridView2.Rows[i].Cells[3].Value == null ||
                       dataGridView2.Rows[i].Cells[4].Value == null ||
                       dataGridView2.Rows[i].Cells[5].Value == null ||
                       dataGridView2.Rows[i].Cells[6].Value == null ||
                       dataGridView2.Rows[i].Cells[7].Value == null
                       )
                    {
                        MessageBox.Show("需求方“" + dataGridView2.Rows[i].Cells[1].Value.ToString() + "”数据缺失！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
                if (
                    dataGridView2.Rows[i].Cells[2].Value != null ||
                    dataGridView2.Rows[i].Cells[3].Value != null ||
                    dataGridView2.Rows[i].Cells[4].Value != null ||
                    dataGridView2.Rows[i].Cells[5].Value != null ||
                    dataGridView2.Rows[i].Cells[6].Value != null ||
                    dataGridView2.Rows[i].Cells[7].Value != null
                    )
                {
                    if (dataGridView2.Rows[i].Cells[1].Value == null)
                    {
                        MessageBox.Show("请补需求方名称！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
            }

            if (CurrentFile == null)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "数据文件（*.csv）|*.csv";
                sfd.RestoreDirectory = true;
                sfd.ShowDialog();
                if(sfd.FileName == "")
                {
                    MessageBox.Show("操作取消！");
                    return;
                }
                SaveFile(sfd.FileName);        
            }
            else
            {
                SaveFile(CurrentFile);
            }
        }

        private void 另存为ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 1000; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value != null)
                {
                    if (
                       dataGridView1.Rows[i].Cells[2].Value == null ||
                       dataGridView1.Rows[i].Cells[3].Value == null ||
                       dataGridView1.Rows[i].Cells[4].Value == null ||
                       dataGridView1.Rows[i].Cells[5].Value == null ||
                       dataGridView1.Rows[i].Cells[6].Value == null ||
                       dataGridView1.Rows[i].Cells[7].Value == null
                       )
                    {
                        MessageBox.Show("供应方“" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "”数据缺失！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
                if (
                    dataGridView1.Rows[i].Cells[2].Value != null ||
                    dataGridView1.Rows[i].Cells[3].Value != null ||
                    dataGridView1.Rows[i].Cells[4].Value != null ||
                    dataGridView1.Rows[i].Cells[5].Value != null ||
                    dataGridView1.Rows[i].Cells[6].Value != null ||
                    dataGridView1.Rows[i].Cells[7].Value != null
                    )
                {
                    if (dataGridView1.Rows[i].Cells[1].Value == null)
                    {
                        MessageBox.Show("请补全供应方名称！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
            }
            for (int i = 0; i < 1000; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value != null)
                {
                    if (
                       dataGridView2.Rows[i].Cells[2].Value == null ||
                       dataGridView2.Rows[i].Cells[3].Value == null ||
                       dataGridView2.Rows[i].Cells[4].Value == null ||
                       dataGridView2.Rows[i].Cells[5].Value == null ||
                       dataGridView2.Rows[i].Cells[6].Value == null ||
                       dataGridView2.Rows[i].Cells[7].Value == null
                       )
                    {
                        MessageBox.Show("需求方“" + dataGridView2.Rows[i].Cells[1].Value.ToString() + "”数据缺失！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
                if (
                    dataGridView2.Rows[i].Cells[2].Value != null ||
                    dataGridView2.Rows[i].Cells[3].Value != null ||
                    dataGridView2.Rows[i].Cells[4].Value != null ||
                    dataGridView2.Rows[i].Cells[5].Value != null ||
                    dataGridView2.Rows[i].Cells[6].Value != null ||
                    dataGridView2.Rows[i].Cells[7].Value != null
                    )
                {
                    if (dataGridView2.Rows[i].Cells[1].Value == null)
                    {
                        MessageBox.Show("请补需求方名称！ 序号 : " + (i + 1).ToString());
                        return;
                    }
                }
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "数据文件（*.csv）|*.csv";
            sfd.RestoreDirectory = true;
            sfd.ShowDialog();
            if (sfd.FileName == "")
            {
                MessageBox.Show("操作取消！");
                return;
            }
            SaveFile(sfd.FileName);
        } 
    }
    /// <summary>
    /// 电厂类，包含电厂提供服务的参数
    /// </summary>
    public class Participator
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string PlantName { get; set; } 
        /// <summary>
        /// 提供/需求低价P1
        /// </summary>
        public double PriceLow { get; set; }
        /// <summary>
        /// 提供/需求中价P2
        /// </summary>
        public double PriceMid { get; set; }
        /// <summary>
        /// 提供/需求高价P3
        /// </summary>
        public double PriceHigh { get; set; }
        /// <summary>
        /// 低价提供/需求量Q1
        /// </summary>
        public double SupplyLow { get; set; }
        /// <summary>
        /// 中价提供/需求量Q2
        /// </summary>
        public double SupplyMid { get; set; }
        /// <summary>
        /// 高价提供/需求量Q3
        /// </summary>
        public double SupplyHigh { get; set; }
        /// <summary>
        /// 生产者成本C
        /// </summary>
        public double TotalCost { get; set; }
        /// <summary>
        ///辅助服务心理价格
        /// </summary>
        public double AimPrice { get; set; }

    }
}
