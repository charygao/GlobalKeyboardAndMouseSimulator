using System;
using System.Threading;
using System.Windows.Forms;

namespace GlobalMacroRecorder
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {

            try
            {
                bool createNew;
                using (new Mutex(true, "Global\\" + Application.ProductName, out createNew))
                {
                    if (createNew)
                    {
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        Application.Run(new MacroForm());
                    }
                    else
                    {
                        MessageBox.Show(@"�����Ѿ������У�");//Only one instance of this application is allowed!
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"�����Ѿ������У�,���г���:" + ex.Message);//Only one instance of this application is allowed!
            }
        }
    }
}
