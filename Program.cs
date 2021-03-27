
using System;
using System.IO;


namespace WindowsShortcuts
{
    struct AFolder
    {

        public AFolder(string path, string icon)
        {
            Path = path;
            Icon = icon;
        }

        public string Path;
        public string Icon;

        public static AFolder[] FromPairs(params string[] pairs)
        {
            var result = new AFolder[pairs.Length / 2];
            for (int i = 0; i < pairs.Length / 2; i++)
            {
                result[i] = new AFolder(
                    pairs[2 * i + 0],
                    pairs[2 * i + 1]
                );
            }

            return result;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var user = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            Console.WriteLine(user);

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            Console.WriteLine(desktop);

            var folders = AFolder.FromPairs(
                @"~\Downloads", "a-downloads",
                @"~\Documents", "a-documents",
                @"~\Projects", "a-projects-git",
                @"~\Google Drive", "a-drive",
                @"~\Google Drive\Projects", "a-projects-drive",
                @"~\Google Drive\Pictures", "a-pictures",
                @"~\Google Drive\Photos\Photography", "a-photography",
                @"~\Google Drive\Photos\Vision", "a-vision"
            );

            var icons = Path.Combine(user, @"Projects\a-windows-shortcuts\icons");


            foreach (var folder in folders)
            {
                string path = folder.Path;

                if (path.StartsWith(@"~\"))
                {
                    path = Path.Combine(user, path.Replace(@"~\", ""));
                }

                string desktopIni = Path.Combine(path, "desktop.ini");

                if (File.Exists(desktopIni))
                {
                    File.Delete(desktopIni);
                }

                string iconLocation = Path.Combine(icons, folder.Icon + ".ico") + ",0";

                try
                {
                    File.WriteAllLines(desktopIni, new String[] {
                        "[.ShellClassInfo]",
                        "IconResource=" + iconLocation,
                        "[ViewState]",
                        "Mode=",
                        "Vid=",
                        "FolderType=Generic"
                    });
                } catch
                {
                    Console.WriteLine("Cant create: desktop.ini");
                }

                IWshRuntimeLibrary.WshShellClass wshShell = new IWshRuntimeLibrary.WshShellClass();

                var link = Path.Combine(desktop, folder.Icon + ".lnk");
                if (File.Exists(link))
                {
                    File.Delete(link);
                }

                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(link);
                shortcut.TargetPath = @"C:\Windows\explorer.exe";
                shortcut.Arguments = path;
                shortcut.IconLocation = iconLocation;
                shortcut.Save();

                Console.WriteLine("Ready: " + path);
            }



            Console.WriteLine("Done: All");
            Console.ReadKey();
        }


    }
}
