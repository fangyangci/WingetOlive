using IWshRuntimeLibrary;
using System.Diagnostics;

class Program
{
    const string AppName = "OliveEnv";
    static string projPath = "";
    static string localAppPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppName);
    static string programAppPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), AppName);

    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            return;
        }

        projPath = args[1];

        if (args[0]== "/Commit")
        {
            Install();
        }
        else if (args[0] == "/Uninstall")
        {
            Uninstall();
        }
    }

    static void Uninstall()
    {
        Console.WriteLine("Python Environment Deleting...");
        RemoveFromEnvPath(Path.Combine(projPath, "OliveScripts"));

        foreach (var subDir in Directory.GetDirectories(projPath))
        {
            Directory.Delete(subDir, true);
        }
        Directory.Delete(programAppPath, true);
        Directory.Delete(localAppPath, true);
    }

    static void Install()
    {
        Console.WriteLine("Python Environment Installing...");
        Console.WriteLine($"Installing Path: {localAppPath}");
        Directory.CreateDirectory(localAppPath);

        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "install.bat";
        // 设置工作目录，即指定在该目录下启动进程
        startInfo.WorkingDirectory = projPath;
        startInfo.ArgumentList.Add(Path.Combine(localAppPath, "Olive"));
        try
        {
            // 使用Process.Start方法启动进程
            var process = Process.Start(startInfo);
            process?.WaitForExit();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"process install error: {ex.Message}");
        }
        Console.WriteLine("Creating StartMenu...");
        try
        {
            CreateStartMenu();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Create StartMenu error: {ex.Message}");
        }

        Console.WriteLine("Install Success!");
    }

    static void CreateStartMenu()
    {
        Console.WriteLine($"StartMenu Path: {programAppPath}");
        Directory.CreateDirectory(programAppPath);
        string shortcutPath = Path.Combine(programAppPath, "Olive Prompt.lnk");

        WshShell shell = new WshShell();
        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

        var cmdPath = Environment.ExpandEnvironmentVariables(@"%WINDIR%\System32\cmd.exe");
        var activatePath = Path.Combine(localAppPath, "Olive", "Scripts", "activate.bat");
        var arguments = $@"/K ""{activatePath}""";

        shortcut.TargetPath = cmdPath;
        shortcut.Arguments = arguments;

        shortcut.WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        shortcut.Description = "Olive Prompt";

        shortcut.IconLocation = Path.Combine(projPath, "OliveMsi", "olive.ico");

        shortcut.Save();

        var olScripts = Path.Combine(projPath, "OliveScripts");
        Directory.CreateDirectory(olScripts);

        System.IO.File.Copy(activatePath, $@"{olScripts}\olivecli.bat", true);
        AddToEnvPath(olScripts);
    }

    static void AddToEnvPath(string addPath)
    {
        var pathKey = "Path";
        var pathStr = Environment.GetEnvironmentVariable(pathKey, EnvironmentVariableTarget.Machine) ?? string.Empty;
        var paths = pathStr.Split(";").ToList();
        if (!paths.Contains(addPath))
        {
            paths.Add(addPath);            
            Environment.SetEnvironmentVariable(pathKey, string.Join(";" ,paths), EnvironmentVariableTarget.Machine);
        }
    }

    static void RemoveFromEnvPath(string remPath)
    {
        var pathKey = "Path";
        var pathStr = Environment.GetEnvironmentVariable(pathKey, EnvironmentVariableTarget.Machine) ?? string.Empty;
        var paths = pathStr.Split(";").ToList();
        if (paths.Contains(remPath))
        {
            paths.Remove(remPath);
            Environment.SetEnvironmentVariable(pathKey, string.Join(";", paths), EnvironmentVariableTarget.Machine);
        }
    }
}