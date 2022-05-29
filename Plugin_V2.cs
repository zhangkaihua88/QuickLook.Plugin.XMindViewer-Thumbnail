using System;
using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Linq;
using QuickLook.Common.Plugin;
using QuickLook.Plugin.ImageViewer;
using QuickLook.Common.Helpers;
using System.Text;
using QuickLook.Plugin.ImageViewer.AnimatedImage;
using QuickLook.Plugin.ImageViewer.AnimatedImage.Providers;
using System.Collections.Generic;

namespace QuickLook.Plugin.XMindViewer
{
    public class Plugin : IViewer
    {
        private ImageViewer.Plugin _viewer;
        private string _imagePath;
        private string logPath;

        private ImagePanel _ip;
        private MetaProvider _meta;

        public int Priority => 0;

        public void Init()
        {
            var useColorProfile = SettingHelper.Get("UseColorProfile", false, "QuickLook.Plugin.ImageViewer");

            AnimatedImage.Providers.Add(
                new KeyValuePair<string[], Type>(
                    useColorProfile ? new[] { ".apng" } : new[] { ".apng", ".png" },
                    typeof(APngProvider)));
            AnimatedImage.Providers.Add(
                new KeyValuePair<string[], Type>(new[] { ".gif" },
                    typeof(GifProvider)));
            AnimatedImage.Providers.Add(
                new KeyValuePair<string[], Type>(
                    useColorProfile ? new string[0] : new[] { ".bmp", ".jpg", ".jpeg", ".jfif", ".tif", ".tiff" },
                    typeof(NativeProvider)));
            AnimatedImage.Providers.Add(
                new KeyValuePair<string[], Type>(new[] { "*" },
                    typeof(ImageMagickProvider)));


            logPath = new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath;
            logPath = Path.GetDirectoryName(logPath);
            // Gets the full path to ensure that relative segments are removed.
            logPath = Path.GetFullPath(Path.Combine(logPath, "log.log"));

            //_viewer = new ImageViewer.Plugin();
            WriteFile(logPath, "Init successful");
            WriteFile(logPath, "==========================");

        }

        public bool CanHandle(string path)
        {
            _imagePath = ExtractFile(path);
            //_viewer.CanHandle(_imagePath);

            WriteFile(logPath, "CanHandle successful");
            WriteFile(logPath, _imagePath);
            WriteFile(logPath, "==========================");
            return !Directory.Exists(path) && new[] { ".xmind" }.Any(path.ToLower().EndsWith);
        }

        public void Prepare(string path, ContextObject context)
        {
            _meta = new MetaProvider(path);
            var size = _meta.GetSize();
            if (!size.IsEmpty)
                context.SetPreferredSizeFit(size, 0.8);
            else
                context.PreferredSize = new Size(800, 600);

            context.Theme = (Themes)SettingHelper.Get("LastTheme", 1, "QuickLook.Plugin.ImageViewer");

            //_viewer.Prepare(_imagePath, context);
            WriteFile(logPath, "Prepare successful");
            WriteFile(logPath, "==========================");
        }

        public void View(string path, ContextObject context)
        {
            _ip = new ImagePanel();
            _ip.ContextObject = context;
            _ip.Meta = _meta;
            _ip.Theme = context.Theme;

            var size = _meta.GetSize();
            context.ViewerContent = _ip;
            context.Title = size.IsEmpty
                ? $"{Path.GetFileName(path)}"
                : $"{size.Width}×{size.Height}: {Path.GetFileName(path)}";

            _ip.ImageUriSource = FilePathToFileUrl(_imagePath);

            //_viewer.View(_imagePath, context);
            WriteFile(logPath, "View successful");
            WriteFile(logPath, context.Title);
            WriteFile(logPath, "==========================");
        }

        public void Cleanup()
        {
            _ip?.Dispose();
            _ip = null;
            WriteFile(logPath, "Cleanup successful");
            WriteFile(logPath, "==========================");
        }

        private string ExtractFile(string path)
        {
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName == "Thumbnails/thumbnail.png")
                    {
                        string curAssemblyFolder = new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath;
                        curAssemblyFolder = Path.GetDirectoryName(curAssemblyFolder);
                        // Gets the full path to ensure that relative segments are removed.
                        string destinationPath = Path.GetFullPath(Path.Combine(curAssemblyFolder, "thumbnail.png"));
                        Console.WriteLine(destinationPath);
                        entry.ExtractToFile(destinationPath, true);
                        return destinationPath;
                    }
                }
                return path;
            }
        }
        private static bool WriteFile(string path, string str)
        {
            try
            {
                StreamWriter sw = File.AppendText(path);
                sw.Write(str);
                sw.Flush();
                sw.Close();
                sw.Dispose();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public static Uri FilePathToFileUrl(string filePath)
        {
            var uri = new StringBuilder();
            foreach (var v in filePath)
                if (v >= 'a' && v <= 'z' || v >= 'A' && v <= 'Z' || v >= '0' && v <= '9' ||
                    v == '+' || v == '/' || v == ':' || v == '.' || v == '-' || v == '_' || v == '~' ||
                    v > '\x80')
                    uri.Append(v);
                else if (v == Path.DirectorySeparatorChar || v == Path.AltDirectorySeparatorChar)
                    uri.Append('/');
                else
                    uri.Append($"%{(int)v:X2}");
            if (uri.Length >= 2 && uri[0] == '/' && uri[1] == '/') // UNC path
                uri.Insert(0, "file:");
            else
                uri.Insert(0, "file:///");

            try
            {
                return new Uri(uri.ToString());
            }
            catch
            {
                return null;
            }
        }
    }
}