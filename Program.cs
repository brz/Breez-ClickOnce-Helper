using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
                var uninstallFlag = false;
                var silentFlag = false;
                var startedFromBrowser = false;
                for (var index = 1; index < commandLineArgs.Length; ++index)
                {
                    if (commandLineArgs[index].Equals("/Uninstall", StringComparison.InvariantCultureIgnoreCase))
                    {
                        uninstallFlag = true;
                    }
                    else if (commandLineArgs[index].Equals("/Silent", StringComparison.InvariantCultureIgnoreCase))
                    {
                        silentFlag = true;
                    }
                }
                if (commandLineArgs.Any(arg => arg.Contains(GetExtensionId(BrowserEnum.MozillaFirefox))) || commandLineArgs.Any(arg => arg.Contains(GetExtensionId(BrowserEnum.GoogleChrome))))
                {
                    startedFromBrowser = true;
                }

                if (startedFromBrowser)
                {
                    var str = OpenStandardStreamIn();

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
                else if (uninstallFlag)
                {
                    Uninstall(silentFlag);
                }
                else
                {
                    Install(silentFlag);
                }
            }
            catch (Exception ex)
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
                Process.Start("rundll32.exe", $"dfshim.dll,ShOpenVerbApplication {url}").WaitForExit();
                return true;
            }
            catch (Exception ex)
            {
#if DEBUG                
                File.AppendAllLines(@"C:\Temp\breezclickoncedebug.txt", new[] { ex.ToString() });
#endif
                return false;
            }
        }

        private static void Install(bool silent)
        {
            try
            {
                //Determine executable path
                var location = Assembly.GetExecutingAssembly().Location;

                //Create target directory
                if (!Directory.Exists(InstallPath))
                {
                    Directory.CreateDirectory(InstallPath);
                }

                //Install breezclickoncehelper.exe
                File.Copy(location, ExecutableFileName, true);

                //Install manifest file for Mozilla Firefox
                var manifestFileNameFirefox = GetManifestFileName(BrowserEnum.MozillaFirefox);
                using (var streamWriter = new StreamWriter(manifestFileNameFirefox, false))
                {
                    streamWriter.Write(GetNativeManifest(BrowserEnum.MozillaFirefox));
                }

                //Install manifest file for Google Chrome
                var manifestFileNameGoogleChrome = GetManifestFileName(BrowserEnum.GoogleChrome);
                using (var streamWriter = new StreamWriter(manifestFileNameGoogleChrome, false))
                {
                    streamWriter.Write(GetNativeManifest(BrowserEnum.GoogleChrome));
                }

                //Registry key for Mozilla Firefox
                Registry.SetValue($@"HKEY_CURRENT_USER\Software\Mozilla\NativeMessagingHosts\{NativeMessageName}", null, manifestFileNameFirefox);

                //Registry key for Google Chrome
                Registry.SetValue($@"HKEY_CURRENT_USER\Software\Google\Chrome\NativeMessagingHosts\{NativeMessageName}", null, manifestFileNameGoogleChrome);

                //Registry key for installed applications
                var keyName = $@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\{ExecutableName}";
                Registry.SetValue(keyName, "Comments", ApplicationDescription);
                Registry.SetValue(keyName, "DisplayName", ApplicationDescription);
                Registry.SetValue(keyName, "DisplayIcon", ExecutableFileName);
                var versionInfo = FileVersionInfo.GetVersionInfo(location);
                Registry.SetValue(keyName, "DisplayVersion", $"{versionInfo.FileMajorPart}.{versionInfo.FileMinorPart}");
                Registry.SetValue(keyName, "Publisher", "Breez");
                Registry.SetValue(keyName, "NoModify", 1);
                Registry.SetValue(keyName, "NoRepair", 1);
                Registry.SetValue(keyName, "UninstallString", $"{ExecutableFileName} /Uninstall");
                Registry.SetValue(keyName, "URLInfoAbout", "https://breezie.be/dev/clickonce");
                if (!silent)
                {
                    MessageBox.Show("Breez ClickOnce Helper was installed successfully", "Information");
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                {
                    MessageBox.Show($"There was an error installing Breez ClickOnce Helper:\n{ex.Message}", "Error");
                }
            }
        }

        private static void Uninstall(bool silent)
        {
            try
            {
                //Remove registry key for Mozilla Firefox
                var subkeyFirefox = $@"Software\Mozilla\NativeMessagingHosts\{NativeMessageName}";
                Registry.CurrentUser.DeleteSubKeyTree(subkeyFirefox);

                //Remove registry key for Google Chrome
                var subkeyChrome = $@"Software\Google\Chrome\NativeMessagingHosts\{NativeMessageName}";
                Registry.CurrentUser.DeleteSubKeyTree(subkeyChrome);

                //Remove registry key for installed applications
                var subkeyUninstall = $@"Software\Microsoft\Windows\CurrentVersion\Uninstall\{ExecutableName}";
                Registry.CurrentUser.DeleteSubKeyTree(subkeyUninstall);

                //Remove manifest file for Mozilla Firefox
                var manifestFileNameFirefox = GetManifestFileName(BrowserEnum.MozillaFirefox);
                if (File.Exists(manifestFileNameFirefox))
                {
                    try
                    {
                        File.Delete(manifestFileNameFirefox);
                    }
                    catch
                    {
                        //Ignore
                    }
                }

                //Remove manifest file for Google Chrome
                var manifestFileNameChrome = GetManifestFileName(BrowserEnum.GoogleChrome);
                if (File.Exists(manifestFileNameChrome))
                {
                    try
                    {
                        File.Delete(manifestFileNameChrome);
                    }
                    catch
                    {
                        //Ignore
                    }
                }

                if (!silent)
                {
                    MessageBox.Show("Breez ClickOnce Helper was uninstalled successfully", "Information");
                }

                //Remove breezclickoncehelper.exe
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
                    MessageBox.Show($"There was an error uninstalling Breez ClickOnce Helper:{ex.Message}", "Error");
                }
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

        private static string GetManifestFileName(BrowserEnum browser)
        {
            return Path.Combine(InstallPath, GetManifestName(browser));
        }

        private const string NativeMessageName = "breez.clickonce.clickoncehelper";

        private const string ApplicationDescription = "Breez ClickOnce Helper";

        private static string GetManifestName(BrowserEnum browser){
            const string ManifestNameFirefox = "breezclickonce_nm_manifest.json";
            const string ManifestNameChrome = "breezclickonce_nm_manifest_chrome.json";
            switch (browser)
            {
                case BrowserEnum.MozillaFirefox:
                    return ManifestNameFirefox;
                case BrowserEnum.GoogleChrome:
                    return ManifestNameChrome;
                default:
                    throw new NotImplementedException();
            }
        }

        private static string GetExtensionId(BrowserEnum browser)
        {
            const string ExtensionIdFirefox = "breezclickonce@breezie.be";
            const string ExtensionIdChrome = "cmkpnegkacdeagdmnbbhjbgdpndmlkgj"; //Store
            //const string ExtensionIdChrome = "dehneackjjjjnhbfemfoehffbaionkcd"; //Debug

            switch (browser)
            {
                case BrowserEnum.MozillaFirefox:
                    return ExtensionIdFirefox;
                case BrowserEnum.GoogleChrome:
                    return ExtensionIdChrome;
                default:
                    throw new NotImplementedException();
            }
        }

        /// <summary>
        /// Gets the json content for a native messaging file for the specified browser.
        /// </summary>
        /// <param name="browser">Mozilla Firefox or Google Chrome</param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private static string GetNativeManifest(BrowserEnum browser)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"{{");
            sb.AppendLine($"  \"name\": \"{NativeMessageName}\",");
            sb.AppendLine($"  \"description\": \"{ApplicationDescription}\",");
            sb.AppendLine($"  \"path\": \"{ExecutableFileName.Replace("\\", "\\\\")}\",");
            sb.AppendLine($"  \"type\": \"stdio\",");
            switch (browser)
            {
                case BrowserEnum.MozillaFirefox:
                    sb.AppendLine($"  \"allowed_extensions\": [\"{GetExtensionId(BrowserEnum.MozillaFirefox)}\"]");
                    break;
                case BrowserEnum.GoogleChrome:
                    sb.AppendLine($"  \"allowed_origins\": [\"chrome-extension://{GetExtensionId(BrowserEnum.GoogleChrome)}/\"]");
                    break;
                default:
                    throw new NotImplementedException();
            }
            sb.AppendLine($"}}");

            return sb.ToString();
        }
    }

    public enum BrowserEnum
    {
        MozillaFirefox = 1,
        GoogleChrome = 2
    }
}
