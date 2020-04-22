using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Web;


namespace Wox.Plugin.OpenCMD
{
    public class Main : IPlugin
    {
        private PluginInitContext context;
        private static List<SystemWindow> openingWindows = new List<SystemWindow>();
        private static string[] names;

        static Main()
        {            
            // use to auto load Interop.SHDocVw.dll from resources
            // only copy to plugin folder can not load correctly
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
        }

        public List<Result> Query(Query query)
        {
            var list = new List<Result>();

            SystemWindow win = null;
            GetOpeningWindows();
            if (openingWindows.Count > 0)
            {
                if (openingWindows[0].Process.ProcessName == "explorer")
                {
                    win = openingWindows[0];
                }
            }

            if (win != null)
            {
                foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindowsClass())
                {
                    var filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                    if (filename.ToLowerInvariant() == "explorer")
                    {
                        if (!window.LocationURL.ToLower().Contains("file:"))
                            continue;

                        // immediately open the windows command
                        if (win.HWnd == (IntPtr)window.HWND)
                        {
                            var path = window.LocationURL.Replace("file:///", "");
                            GitCommandList(list, path);                            
                        }                          
                    }
                }
            }

            if (list.Count <= 0)
            {
                // list all opening folder
                foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindowsClass())
                {
                    var filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                    if (filename.ToLowerInvariant() == "explorer")
                    {
                        if (!window.LocationURL.ToLower().Contains("file:"))
                            continue;

                        var path = window.LocationURL.Replace("file:///", "");
                        path = HttpUtility.UrlDecode(path);
                        if (!Directory.Exists(path))
                            continue;

                        GitCommandList(list, path);                       

                    }
                }
            }
            GitCommandListOfUserPath(list);

            return list;
        }

        private void GitCommandList(List<Result> list, String path)
        {
            String currentDate = System.DateTime.Now.ToString();
            foreach (String cmd in names)
            {
                list.Add(new Result()
                {
                    IcoPath = "Images\\" + cmd + ".png",
                    Title = path,
                    SubTitle = "Open → " + cmd + " ← in this path. " + currentDate,
                    Action = (c) =>
                    {
                        StartShell("-p \"" + cmd + "\" -d \"" + path +"\"");
                        return true;
                    }
                });
            }            
        }

        private void GitCommandListOfUserPath(List<Result> list)
        {
            String currentDate = System.DateTime.Now.ToString();
            foreach (String cmd in names)
            {
                list.Add(new Result()
                {
                    IcoPath = "Images\\" + cmd + ".png",
                    Title = "user default path ~",
                    SubTitle = "Open → " + cmd + " ← in user default path. " + currentDate,
                    Action = (c) =>
                    {
                        StartShell("-p \"" + cmd + "\"");
                        return true;
                    }
                });
            }
            
        }

        public void Init(PluginInitContext context)
        {
            this.context = context;
            String appPath = context.CurrentPluginMetadata.ExecuteFilePath;
            String settingFilePath = appPath.Substring(0, appPath.LastIndexOf("\\")) + "\\setting.info";
            if (System.IO.File.Exists(settingFilePath))
            {
                names = System.IO.File.ReadAllLines(@settingFilePath);
                for (int i = 0; i < names.Length; i++)
                {
                    names[i] = names[i].Trim();
                }
            }
            else
            {
                names = new string[3] { "Windows PowerShell", "cmd", "git", };
            }
        }

        private static void StartShell(string path)
        {
            path = HttpUtility.UrlDecode(path);
            var cmder = Environment.GetEnvironmentVariable("CMDER_ROOT");
            if (!string.IsNullOrEmpty(cmder))
            {
                Process.Start(new ProcessStartInfo()
                {
                    FileName = Path.Combine(cmder, "wt.exe"),
                    Arguments = path
                });
            }
            else
            {
                Process.Start(new ProcessStartInfo()
                {
                    FileName = "cmd",
                    WorkingDirectory = path
                });
            } 
        }

        private static void GetOpeningWindows()
        {
            openingWindows = new List<SystemWindow>();
            WinApi.EnumWindowsProc callback = EnumWindows;
            WinApi.EnumWindows(callback, 0);
        }

        private static bool EnumWindows(IntPtr hWnd, int lParam)
        {
            if (!WinApi.IsWindowVisible(hWnd))
                return true;

            var title = new StringBuilder(256);
            WinApi.GetWindowText(hWnd, title, 256);

            if (string.IsNullOrEmpty(title.ToString()))
            {
                return true;
            }

            if (title.Length != 0 || (title.Length == 0 & hWnd != WinApi.statusbar))
            {
                var window = new SystemWindow(hWnd);
                if (window.IsAltTabWindow() && !window.IsTopmostWindow())
                {
                    openingWindows.Add(window);
                }
            }

            return true;
        }

        public static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string dllName = args.Name.Contains(',') ? args.Name.Substring(0, args.Name.IndexOf(',')) : args.Name.Replace(".dll", "");

            dllName = dllName.Replace(".", "_");

            if (dllName.EndsWith("_resources")) return null;

            System.Resources.ResourceManager rm = new System.Resources.ResourceManager(typeof(Main).Namespace + ".Properties.Resources", System.Reflection.Assembly.GetExecutingAssembly());

            byte[] bytes = (byte[])rm.GetObject(dllName);

            return System.Reflection.Assembly.Load(bytes);
        }


    }
}
