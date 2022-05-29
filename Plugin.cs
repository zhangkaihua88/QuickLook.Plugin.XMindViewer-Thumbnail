using System;
using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Linq;
using QuickLook.Common.Plugin;
using QuickLook.Plugin.ImageViewer;
using System.Text;
using QuickLook.Common.Helpers;

namespace QuickLook.Plugin.XMindViewer
{
    public class Plugin : IViewer
    {
        private string _imagePath;
        private ImagePanel _ip;
        private MetaProvider _meta;
        public int Priority => 0;
        public void Init()
        {
        }

        public bool CanHandle(string path)
        {
            return !Directory.Exists(path) && new[] { ".xmind" }.Any(path.ToLower().EndsWith);
        }

        public void Prepare(string path, ContextObject context)
        {
            _imagePath = ExtractFile(path);
            _meta = new MetaProvider(_imagePath);
            var size = _meta.GetSize();
            if (!size.IsEmpty)
                context.SetPreferredSizeFit(size, 0.8);
            else
                context.PreferredSize = new Size(800, 600);

            context.Theme = (Themes)SettingHelper.Get("LastTheme", 1, "QuickLook.Plugin.ImageViewer");
        }

        public void View(string path, ContextObject context)
        {
            _imagePath = ExtractFile(path);
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
        }

        public void Cleanup()
        {
            _ip?.Dispose();
            _ip = null;
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
                        entry.ExtractToFile(destinationPath, true);
                        return destinationPath;
                    }
                }
                return path;
            }
        }
        public Uri FilePathToFileUrl(string filePath)
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
                return new Uri(filePath);
            }
        }
    }
}