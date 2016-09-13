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
using System.Windows.Shapes;

namespace ProblemSetGenerator
{
    /// <summary>
    /// Interaction logic for GenerateProblem.xaml
    /// </summary>
    public partial class GenerateProblemWindow : Window
    {
        public GenerateProblemWindow()
        {
            InitializeComponent();
            this.GenerateAll.IsChecked = true;
        }

        private void StartGenerating(object sender, RoutedEventArgs e)
        {
            try
            {
                Helper.GenerateProblemSet(MainWindow.workDirectory + "\\题库", this.Problem1.Text + this.Problem2.Text + this.Problem3.Text + this.Problem4.Text);
                MessageBoxResult result = MessageBox.Show("生成完毕。", "成功", MessageBoxButton.OK);
            }
            catch
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。", "错误", MessageBoxButton.OK);
            }
        }

        private void GenerateAll_RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            this.Problem1.IsEnabled = false;
            this.Problem2.IsEnabled = false;
            this.Problem3.IsEnabled = false;
            this.Problem4.IsEnabled = false;
        }

        private void GenerateOne_RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            this.Problem1.IsEnabled = true;
            this.Problem2.IsEnabled = true;
            this.Problem3.IsEnabled = true;
            this.Problem4.IsEnabled = true;
        }
    }
}
