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

namespace ProblemSetGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string workDirectory = "";
        public MainWindow()
        {
            workDirectory = System.IO.Directory.GetCurrentDirectory();

            Helper.CloseWordExcel();
            this.Closing += (object sender, System.ComponentModel.CancelEventArgs e) => { Helper.CloseWordExcel(); };

            InitializeComponent();
            
        }

        private void GenerateProblemSet(object sender, RoutedEventArgs e)
        {
            var window = new GenerateProblemWindow();
            window.Show();
        }

        private void GenerateHomeworkID(object sender, RoutedEventArgs e)
        {
            try
            {
                HomeworkGenerator.GenerateID();
                MessageBoxResult result = MessageBox.Show("生成完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。\n"+ ex.Message, "错误", MessageBoxButton.OK);
            }
        }

        private void GenerateHomeworkByID(object sender, RoutedEventArgs e)
        {
            try
            {
                HomeworkGenerator.GenerateByID(1);
                MessageBoxResult result = MessageBox.Show("生成完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。\n"+ ex.Message, "错误", MessageBoxButton.OK);
            }
        }

        private void GenerateGradeByID(object sender, RoutedEventArgs e)
        {
            try
            {
                HomeworkGenerator.GenerateByID(2);
                MessageBoxResult result = MessageBox.Show("生成完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。\n" + ex.Message, "错误", MessageBoxButton.OK);
            }
        }

        private void Send_Homework_Button_Click(object sender, RoutedEventArgs e)
        {
            var window = new SendEmailWindow();
            window.Show();
        }

        private void Collect_Homework_Button_Click(object sender, RoutedEventArgs e)
        {
            var window = new CollectHomeworkWindow();
            window.Show();
        }

        private void SendFeedback_Button_Click(object sender, RoutedEventArgs e)
        {
            var window = new SendFeedbackWindow();
            window.Show();
        }

        private void Classify_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                HomeworkGenerator.ClassifyHomework();
                MessageBoxResult result = MessageBox.Show("生成完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。\n" + ex.Message, "错误", MessageBoxButton.OK);
            }
            
        }

        private void Merge_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                HomeworkGenerator.MergeFeedback();
                MessageBoxResult result = MessageBox.Show("生成完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。\n" + ex.Message, "错误", MessageBoxButton.OK);
            }

        }
    }
}
