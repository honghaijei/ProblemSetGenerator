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
    /// Interaction logic for CollectHomeworkWindow.xaml
    /// </summary>
    public partial class CollectHomeworkWindow : Window
    {
        public CollectHomeworkWindow()
        {
            InitializeComponent();
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                HomeworkGenerator.CollectHomework(this.EmailBox.Text, this.PasswordBox.Password, this.TitleBox.Text);
                MessageBoxResult result = MessageBox.Show("收集完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。\r\n" + ex.Message , "错误", MessageBoxButton.OK);
            }
        }
    }
}
