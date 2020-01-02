using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.OleDb;//添加类库
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;//添加类库 输入输出
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;//添加类库，输入输出

namespace _20170649_马春龙_5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //角度转弧度
        public double dmstorad(string s)
        {
            string[] ss = s.Split(new char[3] { '°', '′', '″' }, StringSplitOptions.RemoveEmptyEntries);
            double[] d = new double[ss.Length];
            for (int i = 0; i < d.Length; i++)//将度分秒存入双精度数组中
                d[i] = Convert.ToDouble(ss[i]);
            double sign = d[0] >= 0.0 ? 1.0 : -1.0;//判断角度值是否为负值
            double rad = 0;
            if (d.Length == 1)
                rad = Math.Abs(d[0]) * Math.PI / 180;
            else if (d.Length == 2)
                rad = (Math.Abs(d[0]) + d[1] / 60) * Math.PI / 180;
            else
                rad = (Math.Abs(d[0]) + d[1] / 60 + d[2] / 60 / 60) * Math.PI / 180;//将度取绝对值，分化为度，秒化为分
            rad = sign * rad;
            return rad;
        }
        //弧度转角度
        public string radtodms(double rad)
        {
            double sign = rad >= 0.0 ? 1.0 : -1.0;
            rad = Math.Abs(rad) * 180 / Math.PI;//将弧度值取绝对值，并转化度
            double[] d = new double[3];
            d[0] = (int)rad;
            d[1] = (int)((rad - d[0]) * 60);
            d[2] = (rad - d[0] - d[1] / 60) * 60 * 60;//获取秒，不取整
            d[2] = Math.Round(d[2], 2);
            if (d[2] == 60)
            {
                d[1] += 1;
                d[2] -= 60;
                if (d[1] == 60)
                {
                    d[0] += 1;
                    d[1] -= 60;
                }
            }
            d[0] = sign * d[0];//度分秒前加正负号
            string s = Convert.ToString(d[0]) + "°" + Convert.ToString(d[1]) + "′" + Convert.ToString(d[2]) + "″";
            return s;
        }
        public double fangweijiao(double[] sdr, double[] cr)
        {
            double sum = 0;
            for (int i = 1; i < sdr.Length; i++)//从第二行开始循环计算坐标方位角、观测角度累加值
            {
                cr[i] = cr[i - 1] + sdr[i] - Math.PI;//计算坐标方位角/左角
                if (cr[i] >= Math.PI * 2)//判断坐标方位角是否在0到2PI之间
                    cr[i] -= Math.PI * 2;
                else if (cr[i] < 0.0)
                    cr[i] += Math.PI * 2;
                sum += sdr[i];
            }
            return sum;
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null; //清除数据源
            dataGridView1.Rows.Clear(); //清空数据表格的行列
            dataGridView1.Columns.Clear();
            OpenFileDialog file = new OpenFileDialog();
            //声明 打开文件对话框 file
            file.Filter = "Excel文件|*.xls|Excel文件|*.xlsx";
            //文件过滤器，只显示Excel文件
            if (file.ShowDialog() == DialogResult.OK)
            //如果文件正常打开
            {
                string fname = file.FileName; //获取打开的文件名称
                string strSource = "provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source=" + fname + ";Extended Properties='Excel 8.0; HDR=Yes;IMEX=1'"; //准备文件来源信息

                OleDbConnection conn = new OleDbConnection(strSource);
                //Excel文件源放到conn中
                string sqlstring = "SELECT * FROM [Sheet1$]";
                //准备选择表中的Sheet1
                OleDbDataAdapter adapter = new OleDbDataAdapter(sqlstring, conn);
                //声明数据适配器adapter
                DataSet da = new DataSet(); //声明数据集da
                adapter.Fill(da); //使用adapter填充方法
                dataGridView1.DataSource = da.Tables[0];
                //将da.Tables[0]作为dataGridView1的数据源
            }
            else
                return;
        }

        private void txt文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null; //清除数据源
            dataGridView1.Rows.Clear(); //清空数据表格的行
            dataGridView1.Columns.Clear();//清空数据表格的列
            OpenFileDialog file = new OpenFileDialog(); //声明 打开文件类 file
            file.Filter =
            "文本文件|*.txt"; //文件过滤器，只显示txt文件
            if (file.ShowDialog() == DialogResult.OK) //如果文件正常打开
            {
                StreamReader sr = new StreamReader(file.FileName,
                System.Text.Encoding.Default);
                //声明文本读取流，并以文本编码格式读取
                textBox1.Text = sr.ReadToEnd(); //将sr中的内容全部放到textBox中
                sr.Close();
            }
            else
                return;
            string[] str = textBox1.Text.Split(new string[] { "\r\n" },
            StringSplitOptions.RemoveEmptyEntries);
            //将textBox1.Text中按行分割，并放在一维字符串数组中
            string[][] k = new string[str.Length][];
            //定义字符串交错数组，行数与str的长度相同
            for (int i = 0; i < str.Length; i++)
                k[i] = str[i].Split(',');//将str中对应字符串以逗号分割，并放在k中
            dataGridView1.RowCount = k.Length;//定义表格控件的行数，与str长度相同
            dataGridView1.ColumnCount = k[0].Length; //定义表格列数，与k[0]长度相同
            for (int i = 0; i < k[0].Length; i++) //将k中第0行元素放入表格的表头
                dataGridView1.Columns[i].HeaderText = k[0][i];
            for (int i = 1; i < k.Length; i++) //将k中数据元素放入对应的表格中
            {
                for (int j = 0; j < k[i].Length; j++)
                    dataGridView1.Rows[i - 1].Cells[j].Value = k[i][j];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] sd = new string[dataGridView1.RowCount - 5]; //新建一个数组存放观测角度的原始值
            double[] sdr = new double[sd.Length]; //新建一个数组存放观测角度的弧度值
            double[] cr = new double[sd.Length]; //新建一个数组存放计算的坐标方位角
            double sum = 0;
            cr[0] = dmstorad(Convert.ToString(dataGridView1.Rows[0].Cells[4].Value));
            //获取第一个坐标方位角，并将其转换成弧度，放入cr[]数组第一个元素中
            double acd = dmstorad(Convert.ToString
            (dataGridView1.Rows[dataGridView1.RowCount - 6].Cells[4].Value));
            //获取终边坐标方位角，并将其转换成弧度，放入放入acd中用于计算和检核
            for (int i = 1; i < sd.Length; i++) //从第二行开始循环，将观测角度的原始值放入sd[]数组中,并转换成弧度值存放在sdr数组中
            {
                sd[i] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                sdr[i] = dmstorad(sd[i]);
            }
            sum = fangweijiao(sdr, cr); //计算改正前坐标方位角和观测角度总和，分别存储在cr数组和sum中
            dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[1].Value = radtodms(sum);
            //将观测角度总和放入表格中
            double fd, fdx;
            fd = cr[cr.Length - 1] - acd;//计算角度闭合差，单位弧度
            fdx = 60 * Math.Sqrt(sd.Length - 1);//计算角度闭合差限差，单位秒
            dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[1].Value =
            Convert.ToString(Math.Round(fd * 180 / Math.PI * 3600, 2)) + "″";
            //将角度闭合差存入表格中
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value =
            Convert.ToString(Math.Round(fdx, 2)) + "″";//将角度闭合差限差存入表格中
            if (Math.Abs(fd * 180 / Math.PI * 3600) > fdx)//检查角度闭合差是否满足要求
                MessageBox.Show("角度闭合差超限！");
            else
            {
                double vd = -fd / (sd.Length - 1);//分配角度闭合差（观测左角）
                double sumvd = 0;
                for (int i = 1; i < sdr.Length; i++)
                {
                    sdr[i] += vd;//计算改正后的观测角度，并存入sdr数组中
                    sumvd += vd;
                    dataGridView1.Rows[i].Cells[2].Value =
                    Convert.ToString(Math.Round(vd * 180 / Math.PI * 3600, 2)) + "″";
                    //将角度改正数存入表格中
                    dataGridView1.Rows[i].Cells[3].Value = radtodms(sdr[i]);
                }
                if (Math.Round(sumvd, 8) != Math.Round(-fd, 8)) //秒保留2位对应弧度是8位
                    MessageBox.Show("角度改正数分配有误！");
                else
                    dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[2].Value =
                    Convert.ToString(Math.Round(sumvd * 180 / Math.PI * 3600, 2)) + "″";
                //将角度改正数总和存入表格中
                sum = fangweijiao(sdr, cr);//推算改正后的坐标方位角
                if (Math.Round(cr[cr.Length - 1], 8) != Math.Round(acd, 8))
                    MessageBox.Show("坐标方位角推算有误！");
                else
                {
                    dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[3].Value =
                    radtodms(sum); //将改正后观测角度总和放入表格中
                    for (int i = 1; i < cr.Length - 1; i++)//将改正后坐标方位角存入表格
                        dataGridView1.Rows[i].Cells[4].Value = radtodms(cr[i]);
                }
            }
            //至此角度调整和计算完毕
            double x2, y2, x3, y3; //存放已知两个点的x，y坐标
            x2 = Convert.ToDouble(dataGridView1.Rows[1].Cells[12].Value);
            y2 = Convert.ToDouble(dataGridView1.Rows[1].Cells[13].Value);
            x3 = Convert.ToDouble(dataGridView1.Rows[sd.Length - 1].Cells[12].Value);
            y3 = Convert.ToDouble(dataGridView1.Rows[sd.Length - 1].Cells[13].Value);
            double[] sl = new double[sd.Length - 1]; //存放观测距离
            double[] dx = new double[sl.Length]; //存放坐标增量
            double[] dy = new double[sl.Length];
            double suml = 0, sumdx = 0, sumdy = 0;
            for (int i = 1; i < sl.Length; i++)
            {
                sl[i] = Convert.ToDouble(dataGridView1.Rows[i].Cells[5].Value);
                //将观测距离放到sl数组中
                suml += sl[i]; //计算距离总和
                dx[i] = sl[i] * Math.Cos(cr[i]); //利用距离和坐标方位角计算坐标增量
                dy[i] = sl[i] * Math.Sin(cr[i]);
                sumdx += dx[i]; //计算坐标增量总和
                sumdy += dy[i];
            }
            double fx, fy, fxy, k1;
            fx = sumdx - (x3 - x2); //计算坐标增量闭合差
            fy = sumdy - (y3 - y2);
            fxy = Math.Sqrt(fx * fx + fy * fy); //计算导线全长闭合差
            k1 = suml / fxy; //计算导线全长相对闭合差分母
            double[] vx = new double[sl.Length]; //定义数组用于存放坐标增量的改正数及总和
            double[] vy = new double[sl.Length];
            double sumvx = 0, sumvy = 0;
            double[] cx = new double[sl.Length]; //定义数组用于存放改正后的坐标增量及总和
            double[] cy = new double[sl.Length];
            double sumcx = 0, sumcy = 0;
            double[] x = new double[sl.Length + 1]; //定义数组用于存放x，y坐标
            double[] y = new double[sl.Length + 1];
            x[1] = x2;
            y[1] = y2;
            if (k1 < 2000) //判断导线全长相对闭合差是否超限
                MessageBox.Show("导线全长相对闭合差超限！");
            else
            {
                for (int i = 1; i < vx.Length; i++)
                {
                    vx[i] = -fx * sl[i] / suml; //计算坐标增量改正数
                    vy[i] = -fy * sl[i] / suml;
                    sumvx += vx[i]; //计算坐标增量改正数总和
                    sumvy += vy[i];
                }
                if (Math.Round(sumvx, 4) != Math.Round(-fx, 4) || Math.Round(sumvy, 4) != Math.Round(-fy, 4))
                    MessageBox.Show("坐标增量分配有误！");
                else
                {
                    for (int i = 1; i < vx.Length; i++)
                    {
                        cx[i] = dx[i] + vx[i]; //计算改正后坐标增量
                        cy[i] = dy[i] + vy[i];
                        sumcx += cx[i]; //计算改正后坐标增量总和
                        sumcy += cy[i];
                    }
                }
            }
            if (Math.Round(sumcx, 4) != Math.Round(x3 - x2, 4) || Math.Round(sumcy, 4) != Math.Round(y3 - y2, 4))
                MessageBox.Show("改正后的坐标增量计算有误！");
            else
            {
                for (int i = 2; i < x.Length; i++)
                {
                    x[i] = x[i - 1] + cx[i - 1]; //计算x,y坐标
                    y[i] = y[i - 1] + cy[i - 1];
                }
            }
            if (Math.Round(x[x.Length - 1], 4) != Math.Round(x3, 4)
            || Math.Round(y[y.Length - 1], 4) != Math.Round(y3, 4))
                MessageBox.Show("坐标计算有误！");
            else
            {
                for (int i = 1; i < sl.Length; i++)
                {
                    dataGridView1.Rows[i].Cells[6].Value = Convert.ToString(Math.Round(dx[i], 4));//将坐标增量放入表格
                    dataGridView1.Rows[i].Cells[7].Value = Convert.ToString(Math.Round(dy[i], 4));
                    dataGridView1.Rows[i].Cells[8].Value = Convert.ToString(Math.Round(vx[i], 4));//将坐标增量改正数放入表格
                    dataGridView1.Rows[i].Cells[9].Value = Convert.ToString(Math.Round(vy[i], 4));
                    dataGridView1.Rows[i].Cells[10].Value = Convert.ToString(Math.Round(cx[i], 4));
                    dataGridView1.Rows[i].Cells[11].Value = Convert.ToString(Math.Round(cy[i], 4));//将改正后坐标增量放入表格
                    dataGridView1.Rows[i].Cells[12].Value = Convert.ToString(Math.Round(x[i], 3));
                    dataGridView1.Rows[i].Cells[13].Value = Convert.ToString(Math.Round(y[i], 3));//将x,y坐标放入表格
                }
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[5].Value =
                Convert.ToString(Math.Round(suml, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[6].Value =
                Convert.ToString(Math.Round(sumdx, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[7].Value =
                Convert.ToString(Math.Round(sumdy, 4));//将距离总和、坐标增量总和放入表格中
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[8].Value =
                Convert.ToString(Math.Round(sumvx, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[9].Value =
                Convert.ToString(Math.Round(sumvy, 4));//将坐标增量改正数总和、改正后坐标增量总和放入表格中
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[10].Value =
            Convert.ToString(Math.Round(sumcx, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 4].Cells[11].Value =
                Convert.ToString(Math.Round(sumcy, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[7].Value =
                Convert.ToString(Math.Round(fx, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value =
                Convert.ToString(Math.Round(fy, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 3].Cells[10].Value =
                Convert.ToString(Math.Round(fxy, 4));
                dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[11].Value =
                Convert.ToString((int)k1); //导线全长相对闭合差分母取整


            }
        }




        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void 数据导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void excel文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel.Application ex = new Excel.Application(); //声明一个Excel.Application对象 ex
            ex.Visible = true; //使ex可见
            ex.Application.Workbooks.Add(true); //在ex中增加一个工作簿
            for (int i = 0; i < dataGridView1.ColumnCount; i++) //把dataGridView1中的列名存入ex中
            {
                ex.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)//把dataGridView1中的数据存入ex中
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    ex.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
            }

            MessageBox.Show("数据输出已完成!");

        }

        private void txt文件ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string str = ""; //准备一个字符串变量str，用于存放想要输出的内容
            for (int i = 0; i < dataGridView1.ColumnCount; i++) //首先把dataGridView1中的列名存入str中
                str = str + Convert.ToString(dataGridView1.Columns[i].HeaderText).PadRight(21);//该字符串一共占21位，不够的，在右边用空格补齐
            str = str + "\r\n"; //列名保存好后，添加回车换行符
            for (int i = 0; i < dataGridView1.RowCount; i++) //把dataGridView1中的数据存入str中
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    str = str + Convert.ToString(dataGridView1.Rows[i].Cells[j].Value).PadRight(25);
                str = str + "\r\n";//每一行保存好后，添加回车换行符
            } //至此想要输出的内容全部存放在str中
            SaveFileDialog sfile = new SaveFileDialog();//实例化一个保存文件对话框sfile
            sfile.Filter =
            "文本文件|*.txt"; //保存文件类型为txt文件
            if (sfile.ShowDialog() == DialogResult.OK) //如果保存成功
            {
                StreamWriter sw = new StreamWriter(sfile.FileName);
                //实例化文本写入流sw，并用sfile.FileName初始化
                sw.WriteLine(str); //将准备好的str写入sw中
                sw.Close(); //关闭sw
                MessageBox.Show("已导出成功！");
            }
            else return;

        }
    }
}

