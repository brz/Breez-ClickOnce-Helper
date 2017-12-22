using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace BreezClickOnceHelper
{
    internal static class Program
    {
        private static void Main()
        {
            try
            {
                var commandLineArgs = Environment.GetCommandLineArgs();
#if DEBUG
                File.AppendAllLines("C:\\Temp\\breezclickoncedebug.txt", commandLineArgs);
#endif
                var installFlag = false;
                var uninstallFlag = false;
                var silentFlag = false;
                var startedFromBrowser = false;
                for (var index = 1; index < commandLineArgs.Length; ++index)
                {
                    if (commandLineArgs[index].Equals("/Install", StringComparison.InvariantCultureIgnoreCase))
                    {
                        installFlag = true;
                    }
                    else if (commandLineArgs[index].Equals("/Uninstall", StringComparison.InvariantCultureIgnoreCase))
                    {
                        uninstallFlag = true;
                    }
                    else if (commandLineArgs[index].Equals("/Silent", StringComparison.InvariantCultureIgnoreCase))
                    {
                        silentFlag = true;
                    }
                }
                if (commandLineArgs.Length == 3 && commandLineArgs[1] == ManifestFileName && commandLineArgs[2] == ExtensionId)
                {
                    startedFromBrowser = true;
                }

                if (installFlag)
                {
                    Install(silentFlag);
                }
                else if (uninstallFlag)
                {
                    Uninstall(silentFlag);
                }
                else
                {
                    var str = OpenStandardStreamIn();
                    if (string.IsNullOrWhiteSpace(str))
                    {
                        if (!IsInstalled)
                        {
                            Install(silentFlag);
                        }
                        else
                        {
                            Uninstall(silentFlag);
                        }
                    }
                    else
                    {
                        if (startedFromBrowser)
                        {
                            //Check installation status
                            if (str.StartsWith("{\"message\":\"", StringComparison.InvariantCultureIgnoreCase) && str.EndsWith("\"}", StringComparison.InvariantCultureIgnoreCase))
                            {
                                var message = str.Substring("{\"message\":\"".Length, str.Length - "{\"message\":\"".Length - "\"}".Length);
                                if (message == "CheckStatus")
                                {
                                    OpenStandardStreamOut("{\"result\":\"OK\"}");
                                }
                            }
                            //Check for ClickOnce URL
                            else if (str.StartsWith("{\"url\":\"", StringComparison.InvariantCultureIgnoreCase) && str.EndsWith("\"}", StringComparison.InvariantCultureIgnoreCase))
                            {
                                var launchOk = LaunchClickOnce(str.Substring("{\"url\":\"".Length, str.Length - "{\"url\":\"".Length - "\"}".Length));

                                if (launchOk)
                                {
                                    OpenStandardStreamOut("{\"result\":\"OK\"}");
                                }
                                else
                                {
                                    OpenStandardStreamOut("{\"result\":\"FAIL\"}");
                                }
                            }
                                
                        }
                    }
                }
            }
            catch(Exception ex)
            {
#if DEBUG
                File.AppendAllLines("C:\\Temp\\breezclickoncedebug.txt", new[] { ex.ToString() });
#endif
            }
        }

        private static string OpenStandardStreamIn()
        {
            var sb = new StringBuilder();
            using (var stream = Console.OpenStandardInput())
            {
                var buffer = new byte[4];
                stream.Read(buffer, 0, 4);
                var int32 = BitConverter.ToInt32(buffer, 0);
                for (var index = 0; index < int32; ++index)
                {
                    sb.Append(char.ToString((char)stream.ReadByte()));
                }
            }

            var pathToLaunch = sb.ToString().TrimStart('"').TrimEnd('"');

#if DEBUG
            File.AppendAllLines("C:\\Temp\\breezclickoncedebug.txt", new[] { "INPUT:", pathToLaunch });
#endif

            return pathToLaunch;
        }

        private static void OpenStandardStreamOut(string stringData)
        {
            var bytes = Encoding.UTF8.GetBytes(stringData);

            var stdout = Console.OpenStandardOutput();
            stdout.WriteByte((byte)((bytes.Length >> 0) & 0xFF));
            stdout.WriteByte((byte)((bytes.Length >> 8) & 0xFF));
            stdout.WriteByte((byte)((bytes.Length >> 16) & 0xFF));
            stdout.WriteByte((byte)((bytes.Length >> 24) & 0xFF));
            stdout.Write(bytes, 0, bytes.Length);
            stdout.Flush();
        }

        private static bool LaunchClickOnce(string url)
        {
            try
            {
                Process.Start("rundll32.exe", "dfshim.dll,ShOpenVerbApplication " + url).WaitForExit();
                return true;
            }
            catch(Exception ex)
            {
#if DEBUG                
                File.AppendAllLines("C:\\Temp\\breezclickoncedebug.txt", new[] { ex.ToString() });
#endif
                return false;
            }
        }

        private static void Install(bool silent)
        {
            try
            {
                var location = Assembly.GetExecutingAssembly().Location;
                
                var str2 = string.Format(string.Join("\n", NativeManifest));
                Directory.CreateDirectory(InstallPath);
                File.Copy(location, ExecutableFileName, true);
                using (var streamWriter = new StreamWriter(ManifestFileName, false))
                {
                    streamWriter.Write(str2);
                }
                Registry.SetValue($"HKEY_CURRENT_USER\\Software\\Mozilla\\NativeMessagingHosts\\{NativeMessageName}", null, ManifestFileName);
                var keyName = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\" + ExecutableName;
                Registry.SetValue(keyName, "Comments", ApplicationDescription);
                Registry.SetValue(keyName, "DisplayName", ApplicationDescription);
                Registry.SetValue(keyName, "DisplayIcon", ExecutableFileName);
                var versionInfo = FileVersionInfo.GetVersionInfo(location);
                Registry.SetValue(keyName, "DisplayVersion", $"{versionInfo.FileMajorPart}.{versionInfo.FileMinorPart}");
                Registry.SetValue(keyName, "Publisher", "Breez");
                Registry.SetValue(keyName, "NoModify", 1);
                Registry.SetValue(keyName, "NoRepair", 1);
                Registry.SetValue(keyName, "UninstallString", ExecutableFileName + " /Uninstall");
                Registry.SetValue(keyName, "URLInfoAbout", "http://breezie.be");
                if (!silent)
                {
                    MessageBox.Show("Breez ClickOnce Helper was installed successfully", "Information");
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                {
                    MessageBox.Show("There was an error installing Breez ClickOnce Helper:\n" + ex.Message, "Error");
                }
            }
        }

        private static void Uninstall(bool silent)
        {
            try
            {
                var subkey1 = $"Software\\Mozilla\\NativeMessagingHosts\\{NativeMessageName}";
                Registry.CurrentUser.DeleteSubKeyTree(subkey1);
                var subkey2 = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\" + ExecutableName;
                Registry.CurrentUser.DeleteSubKeyTree(subkey2);
                File.Delete(ManifestFileName);
                if (!silent)
                {
                    MessageBox.Show("Breez ClickOnce Helper was uninstalled successfully", "Information");
                }
                Process.Start(new ProcessStartInfo()
                {
                    Arguments = $"/C choice /C Y /N /D Y /T 3 & rd /S /Q \"{InstallPath}\"",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    FileName = "cmd.exe"
                });
            }
            catch (Exception ex)
            {
                if (!silent)
                {
                    MessageBox.Show("There was an error uninstalling Breez ClickOnce Helper:\n" + ex.Message, "Error");
                }
            }
        }
        
        #region Consts and Properties
        private static bool IsInstalled
        {
            get
            {
                var installed = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ExecutableName, false) != null;
                return installed;
            }
        }

        private static string InstallPath
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Breez", "ClickOnceHelper");
            }
        }

        private static string ExecutableName
        {
            get
            {
                return Assembly.GetExecutingAssembly().ManifestModule.ScopeName;
            }
        }

        private static string ExecutableFileName
        {
            get { return Path.Combine(InstallPath, ExecutableName); }
        }

        private static string ManifestFileName
        {
            get { return Path.Combine(InstallPath, ManifestName); }
        }

        private const string ManifestName = "breezclickonce_nm_manifest.json";

        private const string ExtensionId = "breezclickonce@breezie.be";

        private const string NativeMessageName = "breez.clickonce.clickoncehelper";

        private const string ApplicationDescription = "Breez ClickOnce Helper";

        private static readonly string[] NativeManifest = {
            "{{",
            $"  \"name\": \"{NativeMessageName}\",",
            $"  \"description\": \"{ApplicationDescription}\",",
            $"  \"path\": \"{ExecutableFileName.Replace("\\", "\\\\")}\",",
            "  \"type\": \"stdio\",",
            $"  \"allowed_extensions\": [\"{ExtensionId}\"]",
            "}}"
        };
        #endregion
    }
}
