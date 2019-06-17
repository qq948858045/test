using Microsoft.Win32;
using RuiYa.Common;
using RuiYa.DataAccess;
using RuiYa.WPF.DialogWindow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Visifire.Charts;

namespace RuiYa.WPF
{
    /// <summary>
    /// CarbonDioxideEmissions.xaml 的交互逻辑
    /// </summary>
    public partial class CarbonDioxideEmissions : Page
    {
        public Model.pro pros = null;
        public Model.maplist maplist = null;
        public Model.sol sols = null;
        private IDAO dao = new DAO();
        List<Model.constant> constanlist = new List<Model.constant>();
        private MainLoad _parentWinLoad;
        public MainLoad ParentWindowLoad
        {
            get { return _parentWinLoad; }
            set { _parentWinLoad = value; }
        }

        public CarbonDioxideEmissions()
        {
            InitializeComponent();
            this.Loaded += CarbonDioxideEmissions_Loaded;
        }

        private void CarbonDioxideEmissions_Loaded(object sender, RoutedEventArgs e)
        {
            constanlist = dao.CurrentDBContext.constants.ToList();
            string[] strlist = new string[] { "吸收碳的森林面积（公顷）", "未消耗的汽油（升）", "未用小汽车小卡车（辆）" };
            comType.ItemsSource = strlist;
            CreateChartSpline();
            comType.SelectedIndex = 0;
            txtVersion.Text = Math.Round(Benchmark - Optimization, 1) + "吨二氧化碳等于" + Math.Round((Benchmark - Optimization) / 168, 1);
        }

        double Benchmark = 0;
        double Optimization = 0;

        public void CreateChartSpline()
        {
            if (pros == null) return;
            if (maplist == null) return;
            if (sols == null) return;
            Calculation ca = new Calculation();
            ca.pros = pros;
            ca.maplist = maplist;
            ca.sols = sols;
            List<DeliverabilityDaily> listDd = new List<DeliverabilityDaily>();
            winProgressBar wp = new winProgressBar();
            var enconfig = dao.CurrentDBContext.energyconfigs.Where(c => c.ProID == pros.id).ToList();
            var coals = enconfig.Where(c => c.EnergyType == 7).FirstOrDefault();
            var naturalgases = enconfig.Where(c => c.EnergyType == 8).FirstOrDefault();
            var dust = Convert.ToDouble(constanlist.Where(c => c.Name == "dust").FirstOrDefault().Value);
            var SO2 = Convert.ToDouble(constanlist.Where(c => c.Name == "SO2").FirstOrDefault().Value);
            var NOX = Convert.ToDouble(constanlist.Where(c => c.Name == "NOX").FirstOrDefault().Value);
            wp.pb.Minimum = 0;
            wp.pb.Value = 0;
            wp.txtMessage.Text = "正在查询能源数据....";
            wp.pb.Maximum = 2;
            wp.Owner = Application.Current.MainWindow;
            wp.ShowInTaskbar = false;
            wp.Show();
            listDd = ca.GetCalculation1(wp);
            if (listDd.Count == 0)
            {
                wp.Close();
                return;
            }
            txtModelName.Text = pros.name;
            txtSolName.Text = sols.Name;
            txtTime.Text = ca.Start.ToShortDateString() + "-" + ca.End.ToShortDateString();
            Chart chart = new Chart();
            GridChart.Children.Add(chart);
            var Canuse = listDd.Sum(c => c.Canuse);
            var Coal = listDd.Sum(c => c.Coal);
            var NaturalGas = listDd.Sum(c => c.NaturalGas);
            #region 图表初始化设置          
            chart.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            chart.VerticalAlignment = System.Windows.VerticalAlignment.Stretch;
            chart.Margin = new Thickness(0, 100, 0, 0);
            //是否启用打印和保持图片
            chart.ToolBarEnabled = false;
            //设置图标的属性
            chart.ScrollingEnabled = false;//是否启用或禁用滚动
            chart.View3D = false;//3D效果显示
            //创建一个标题的对象
            Title title = new Title();

            //设置标题的名称
            title.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chart.Titles.Add(title);
            chart.Titles[0].FontSize = 25;
            //chart.Titles[0].FontColor = Brushes.Blue;
            chart.LightingEnabled = true;


            //初始化一个新的Axis
            Axis xaxis = new Axis();
            //设置Axis的属性
            //图表的X轴坐标按什么来分类，如时分秒

            //图表的X轴坐标间隔如2,3,20等，单位为xAxis.IntervalType设置的时分秒。
            xaxis.Interval = 1;
            //设置X轴的时间显示格式为7-10 11：20   
            xaxis.IntervalType = IntervalTypes.Auto;
            AxisLabels xal = new AxisLabels();
            xaxis.AxisLabels = xal;
            //给图标添加Axis     
            chart.AxesX.Add(xaxis);
            Axis yAxis = new Axis();
            //设置图标中Y轴的最小值永远为0           
            //yAxis.AxisMinimum = 0;
            //设置图表中Y轴的后缀          
            yAxis.Suffix = "吨";
            chart.AxesY.Add(yAxis);
            #endregion

            #region 两条数据线操作
            // 创建产能数据线。               
            DataSeries dataSeries = new DataSeries();
            // 设置数据线的格式。               
            dataSeries.LegendText = "基准方案排放量（吨）";
            dataSeries.Color = new SolidColorBrush(Color.FromRgb(255, 153, 0));
            dataSeries.RenderAs = RenderAs.Column;//柱状图
            dataSeries.XValueType = ChartValueTypes.Auto;
            dataSeries.Width = 20;
            chart.Series.Add(dataSeries);
            // 创建用能数据线。   
            DataSeries dataSeries2 = new DataSeries();
            // 设置数据线的格式。         
            dataSeries2.LegendText = "优化方案放量（吨）";
            dataSeries2.Color = new SolidColorBrush(Color.FromRgb(51, 204, 102));
            dataSeries2.RenderAs = RenderAs.Column;//柱状图
            dataSeries2.XValueType = ChartValueTypes.Auto;
            chart.Series.Add(dataSeries2);
            #endregion
            for (int i = 0; i < 4; i++)
            {

                //wp.DoEvents();
                // 创建一个数据点的实例。                   
                DataPoint dataPoint = new DataPoint();
                // 设置X轴点                    
                dataPoint.XValue = i;
                //设置Y轴点                   
                dataPoint.MarkerSize = 2;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPoint.Color = new SolidColorBrush(Color.FromRgb(255, 153, 0));
                //添加数据点                   
                //wp.DoEvents();
                // 创建一个数据点的实例。                   
                DataPoint dataPoint1 = new DataPoint();
                // 设置X轴点                    
                dataPoint1.XValue = i;
                //设置Y轴点                   
                dataPoint1.MarkerSize = 2;
                //dataPoint.Tag = tableName.Split('(')[0];
                //设置数据点颜色                  
                dataPoint1.Color = new SolidColorBrush(Color.FromRgb(51, 204, 102));
                //添加数据点                   
                if (i == 0)
                {
                    dataPoint.AxisXLabel = "二氧化碳";
                    dataPoint1.AxisXLabel = "二氧化碳";
                    Benchmark = dataPoint.YValue = Convert.ToDouble(Canuse * coals.Rhl);
                    Optimization = dataPoint1.YValue = Convert.ToDouble(Coal * coals.Rhl + NaturalGas * naturalgases.Rhl);
                    txtBenchmark1.Text = dataPoint.YValue.ToString();
                    txtOptimization1.Text = dataPoint1.YValue.ToString();
                    txtEmission1.Text = ((dataPoint.YValue - dataPoint1.YValue) / dataPoint.YValue * 100).ToString();
                    txrEmission.Text = Math.Round(Convert.ToDouble(txtBenchmark1.Text) - Convert.ToDouble(txtOptimization1.Text), 2).ToString();
                }
                else if (i == 1)
                {
                    dataPoint.AxisXLabel = "碳粉尘";
                    dataPoint1.AxisXLabel = "碳粉尘";
                    dataPoint.YValue = Convert.ToDouble(Canuse) * dust;
                    dataPoint1.YValue = Convert.ToDouble(Coal) * dust;
                    txtBenchmark2.Text = dataPoint.YValue.ToString();
                    txtOptimization2.Text = dataPoint1.YValue.ToString();
                    txtEmission2.Text = ((dataPoint.YValue - dataPoint1.YValue) / dataPoint.YValue * 100).ToString();
                }
                else if (i == 2)
                {
                    dataPoint.AxisXLabel = "二氧化硫";
                    dataPoint1.AxisXLabel = "二氧化硫";
                    dataPoint.YValue = Convert.ToDouble(Canuse) * SO2;
                    dataPoint1.YValue = Convert.ToDouble(Coal) * SO2;
                    txtBenchmark3.Text = dataPoint.YValue.ToString();
                    txtOptimization3.Text = dataPoint1.YValue.ToString();
                    txtEmission3.Text = ((dataPoint.YValue - dataPoint1.YValue) / dataPoint.YValue * 100).ToString();
                }
                else if (i == 3)
                {
                    dataPoint.AxisXLabel = "氮氧化物";
                    dataPoint1.AxisXLabel = "氮氧化物";
                    dataPoint.YValue = Convert.ToDouble(Canuse) * NOX;
                    dataPoint1.YValue = Convert.ToDouble(Coal) * NOX;
                    txtBenchmark4.Text = dataPoint.YValue.ToString();
                    txtOptimization4.Text = dataPoint1.YValue.ToString();
                    txtEmission4.Text = ((dataPoint.YValue - dataPoint1.YValue) / dataPoint.YValue * 100).ToString();
                }

                dataSeries.DataPoints.Add(dataPoint);
                dataSeries2.DataPoints.Add(dataPoint1);
            }
            wp.Close();
            GridTable.Visibility = Visibility.Visible;
            GridIcon.Visibility = Visibility.Visible;
            GridIncome.Visibility = Visibility.Visible;
        }

        private void winModelSol_Click(object sender, RoutedEventArgs e)
        {
            winModelSol ms = new winModelSol();
            ms.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ms.Owner = Application.Current.MainWindow;
            ms.ShowInTaskbar = false;
            if (ms.ShowDialog() == true)
            {
                maplist = ms.maplist;
                pros = ms.pros;
                sols = ms.sols;
                if (sols == null)
                {
                    btnRun.Visibility = Visibility.Hidden;
                    btnExetb.Visibility = Visibility.Visible;
                    GridTable.Visibility = Visibility.Hidden;
                    GridIcon.Visibility = Visibility.Hidden;
                    GridIncome.Visibility = Visibility.Hidden;
                    btnRun1.Visibility = Visibility.Visible;
                    Grid1.Visibility = Visibility.Visible;
                }
                else
                {
                    btnRun.Visibility = Visibility.Visible;
                    btnExetb.Visibility = Visibility.Hidden;
                    GridTable.Visibility = Visibility.Visible;
                    GridIcon.Visibility = Visibility.Visible;
                    GridIncome.Visibility = Visibility.Visible;
                    btnRun1.Visibility = Visibility.Hidden;
                    Grid1.Visibility = Visibility.Hidden;
                }
                GridChart.Children.Clear();
                btnRun.Style = FindResource("Mybutton62") as Style;
                btnRun.IsEnabled = true;
            }
        }

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            CreateChartSpline();
            comType_Selected(sender, e);
        }

        private void comType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comType.SelectedIndex == 0)
            {
                txtVersion.Text = Math.Round(Benchmark - Optimization, 1) + "吨二氧化碳等于" + Math.Round((Benchmark - Optimization) / 168, 1);
                icon.Source = new BitmapImage(new Uri("pack://application:,,,/RuiYa.WPF;component/images/公顷森林.png"));
            }
            else if (comType.SelectedIndex == 1)
            {
                txtVersion.Text = Math.Round(Benchmark - Optimization, 1) + "吨二氧化碳等于" + Math.Round((Benchmark - Optimization) * 1000 / 2.361, 1);
                icon.Source = new BitmapImage(new Uri("pack://application:,,,/RuiYa.WPF;component/images/汽油.png"));
            }
            else if (comType.SelectedIndex == 2)
            {
                txtVersion.Text = Math.Round(Benchmark - Optimization, 1) + "吨二氧化碳等于" + Math.Round((Benchmark - Optimization) / 5.5, 1);
                icon.Source = new BitmapImage(new Uri("pack://application:,,,/RuiYa.WPF;component/images/汽车.png"));
            }
            else
            {
                txtVersion.Text = "0.0吨二氧化碳等于0.0";
            }
        }

        private void comType_Selected(object sender, RoutedEventArgs e)
        {
            if (comType.SelectedIndex == 0)
            {
                txtVersion.Text = Math.Round(Benchmark - Optimization, 1) + "吨二氧化碳等于" + Math.Round((Benchmark - Optimization) / 168, 1);
                icon.Source = new BitmapImage(new Uri("pack://application:,,,/RuiYa.WPF;component/images/公顷森林.png"));
            }
            else
            {
                txtVersion.Text = "0.0吨二氧化碳等于0.0";
            }
        }

        private void RunBtn1_Click(object sender, RoutedEventArgs e)
        {
            GetModelRun();
            btnExetb.Style = FindResource("Mybutton66") as Style;
            btnExetb.IsEnabled = true;
        }

        Calculation ca = new Calculation();
        List<Model.DischargeClass> disList = new List<Model.DischargeClass>();
        public void GetModelRun()
        {
            if (pros == null) return;
            if (maplist == null) return;
            var solList = dao.CurrentDBContext.sols.Where(c => c.ProID == pros.id).ToList();


            ca.pros = pros;
            ca.maplist = maplist;
            foreach (var item in solList)
            {
                ca.sols = item;
                Model.DischargeClass dis = new Model.DischargeClass();
                dis.Name = item.Name;
                dis.User = item.User;
                List<DeliverabilityDaily> listDd = new List<DeliverabilityDaily>();
                winProgressBar wp = new winProgressBar();
                var enconfig = dao.CurrentDBContext.energyconfigs.Where(c => c.ProID == pros.id).ToList();
                var coals = enconfig.Where(c => c.EnergyType == 7).FirstOrDefault();
                var naturalgases = enconfig.Where(c => c.EnergyType == 8).FirstOrDefault();
                var dust = Convert.ToDecimal(constanlist.Where(c => c.Name == "dust").FirstOrDefault().Value);
                var SO2 = Convert.ToDecimal(constanlist.Where(c => c.Name == "SO2").FirstOrDefault().Value);
                var NOX = Convert.ToDecimal(constanlist.Where(c => c.Name == "NOX").FirstOrDefault().Value);
                wp.pb.Minimum = 0;
                wp.pb.Value = 0;
                wp.txtMessage.Text = "正在查询能源数据....";
                wp.pb.Maximum = 2;
                wp.Owner = Application.Current.MainWindow;
                wp.ShowInTaskbar = false;
                wp.Show();
                listDd = ca.GetCalculation1(wp);
                if (listDd.Count == 0)
                {
                    wp.Close();
                    return;
                }
                wp.Close();
                var Canuse = listDd.Sum(c => c.Canuse);
                var Coal = listDd.Sum(c => c.Coal);
                var NaturalGas = listDd.Sum(c => c.NaturalGas);

                dis.Benchmark1 = (decimal)Math.Round(Convert.ToDouble(Canuse * coals.Rhl), 2);
                dis.Optimization1 = (decimal)Math.Round(Convert.ToDouble(Coal * coals.Rhl + NaturalGas * naturalgases.Rhl), 2);
                dis.Emission1 = (decimal)Math.Round(Convert.ToDouble((dis.Benchmark1 - dis.Optimization1) / dis.Benchmark1 * 100), 2);

                dis.Benchmark2 = (decimal)Math.Round(Convert.ToDouble(Canuse * dust), 2);
                dis.Optimization2 = (decimal)Math.Round(Convert.ToDouble(Coal * dust), 2);
                dis.Emission2 = (decimal)Math.Round(Convert.ToDouble((dis.Benchmark2 - dis.Optimization2) / dis.Benchmark2 * 100), 2);

                dis.Benchmark3 = (decimal)Math.Round(Convert.ToDouble(Canuse * SO2), 2);
                dis.Optimization3 = (decimal)Math.Round(Convert.ToDouble(Coal * SO2), 2);
                dis.Emission3 = (decimal)Math.Round(Convert.ToDouble((dis.Benchmark3 - dis.Optimization3) / dis.Benchmark3 * 100), 2);

                dis.Benchmark4 = (decimal)Math.Round(Convert.ToDouble(Canuse * NOX), 2);
                dis.Optimization4 = (decimal)Math.Round(Convert.ToDouble(Coal * NOX), 2);
                dis.Emission4 = (decimal)Math.Round(Convert.ToDouble((dis.Benchmark4 - dis.Optimization4) / dis.Benchmark4 * 100), 2);
                disList.Add(dis);
            }
            Grid1.ItemsSource = disList;
        }

        private void txtPay_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPay.Text) || string.IsNullOrWhiteSpace(txrEmission.Text)) return;
            txtIncome.Text = (Convert.ToDecimal(txtPay.Text) * Convert.ToDecimal(txrEmission.Text) / 10000).ToString();
        }

        private void btnExeTable(object sender, RoutedEventArgs e)
        {
            if (pros == null) return;
            if (maplist == null) return;
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel文件|*.xls;*.xlsx";
            if (dialog.ShowDialog() == false)
            {
                return;
            }
            winProgressBar wp = new winProgressBar();
            wp.pb.Minimum = 0;
            wp.pb.Value = 0;
            wp.txtMessage.FontSize = 12;
            wp.txtMessage.Text = "正在导出，请稍后....";
            wp.pb.Maximum = 8;
            wp.Owner = Application.Current.MainWindow;
            wp.ShowInTaskbar = false;
            wp.Show();
            wp.StepAdd(1);
            var strExcelFileName = Environment.CurrentDirectory + "/排放导出标准文档.xls";
            ExcelHelper expHelper = new ExcelHelper(strExcelFileName);
            if (string.IsNullOrEmpty(expHelper.errorMsg) == false)
            {
                winMessage wm1 = new winMessage();
                wm1.txtMessage.Text = "打开" + strExcelFileName + "文件有错误：" + expHelper.errorMsg;
                wm1.Title.Text = "提示";
                wm1.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                wm1.Owner = Application.Current.MainWindow;
                wm1.ShowInTaskbar = false;
                wm1.ShowDialog();
                return;
            }

            //第一个表综合月报表
            int row = 2;
            int ESIndex = 0;
            int i = 1;
            foreach (var item in disList)
            {
                wp.DoEvents();
                if (wp.isCancel)
                {
                    expHelper.SaveAs(dialog.FileName);
                    expHelper.Close();
                    return;
                }
                wp.txtMessage.Text = string.Format("正在导出，请稍后....", row);
                expHelper.setCellValue(row, 0, i, ESIndex);
                expHelper.setCellValue(row, 1, item.Name, ESIndex);
                expHelper.setCellValue(row, 2, item.User, ESIndex);
                expHelper.setCellValue(row, 3, item.Benchmark1, ESIndex);
                expHelper.setCellValue(row, 4, item.Optimization1, ESIndex);
                expHelper.setCellValue(row, 5, item.Emission1, ESIndex);
                expHelper.setCellValue(row, 6, item.Benchmark2, ESIndex);
                expHelper.setCellValue(row, 7, item.Optimization2, ESIndex);
                expHelper.setCellValue(row, 8, item.Emission2, ESIndex);
                expHelper.setCellValue(row, 9, item.Benchmark3, ESIndex);
                expHelper.setCellValue(row, 10, item.Optimization3, ESIndex);
                expHelper.setCellValue(row, 11, item.Emission3, ESIndex);
                expHelper.setCellValue(row, 12, item.Benchmark4, ESIndex);
                expHelper.setCellValue(row, 13, item.Optimization4, ESIndex);
                expHelper.setCellValue(row, 14, item.Emission4, ESIndex);
                i++;
            }
            wp.StepAdd(1);
            wp.txtMessage.Text = "成功导出。";
            wp.Close();
            expHelper.SaveAs(dialog.FileName);
            expHelper.Close();

            winMessage wm2 = new winMessage();
            wm2.txtMessage.Text = "是否打开Excel文件";
            wm2.Title.Text = "打开询问";
            wm2.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            wm2.Owner = Application.Current.MainWindow;
            wm2.ShowInTaskbar = false;
            if (wm2.ShowDialog() == true)
            {
                System.Diagnostics.Process.Start(dialog.FileName);
            }
        }
    }
}
