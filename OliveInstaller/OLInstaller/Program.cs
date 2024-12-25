using IWshRuntimeLibrary;
using System.Diagnostics;

class Program
{
    static string projPath = "";
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
        foreach (var subDir in Directory.GetDirectories(projPath))
        {
            Directory.Delete(subDir, true);
        }
        RemoveFromPath(Path.Combine(projPath, "OliveScripts"));

        Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "OliveEnv"), true);
    }

    static void Install()
    {
        Console.WriteLine("Python Environment Installing...");
        Console.WriteLine($"Installing Path: {projPath}");

        ProcessStartInfo startInfo = new ProcessStartInfo();
        startInfo.FileName = "install.bat";
        // 设置工作目录，即指定在该目录下启动进程
        startInfo.WorkingDirectory = projPath;
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
        string startMenuFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "OliveEnv");
        Directory.CreateDirectory(startMenuFolder);
        string shortcutPath = Path.Combine(startMenuFolder, "Olive Prompt.lnk");

        WshShell shell = new WshShell();
        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

        var cmdPath = Environment.ExpandEnvironmentVariables(@"%WINDIR%\System32\cmd.exe");
        var activatePath = Path.Combine(projPath, "Olive", "Scripts", "activate.bat");
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
        AddToPath(olScripts);
    }

    static void AddToPath(string addPath)
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

    static void RemoveFromPath(string remPath)
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