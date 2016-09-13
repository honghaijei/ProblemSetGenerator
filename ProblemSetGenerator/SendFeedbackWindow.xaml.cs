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
    /// Interaction logic for SendFeedbackWindow.xaml
    /// </summary>
    public partial class SendFeedbackWindow : Window
    {
        public SendFeedbackWindow()
        {
            InitializeComponent();
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string content = this.ContentTextBox.Text;
                HomeworkGenerator.SendFeedback(this.EmailBox.Text, this.Account.Text, this.PasswordBox.Password, this.TitleBox.Text, content, this.SMTPAddress.Text, int.Parse(this.SMTPPort.Text));
                MessageBoxResult result = MessageBox.Show("发送完毕。", "成功", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBoxResult result = MessageBox.Show("请检查配置是否正确。" + ex.Message, "错误", MessageBoxButton.OK);
            }
        }
    }
}
