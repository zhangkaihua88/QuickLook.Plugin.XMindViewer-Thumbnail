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
            logPath = new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath;
            logPath = Path.GetDirectoryName(logPath);
            // Gets the full path to ensure that relative segments are removed.
            logPath = Path.GetFullPath(Path.Combine(logPath, "log.log"));

            _viewer = new ImageViewer.Plugin();
            WriteFile(logPath, "Init successful");
            WriteFile(logPath, "==========================");

        }

        public bool CanHandle(string path)
        {
            _imagePath = ExtractFile(path);
            _viewer.CanHandle(_imagePath);

            WriteFile(logPath, "CanHandle successful");
            WriteFile(logPath, _imagePath);
            WriteFile(logPath, "==========================");
            return !Directory.Exists(path) && new[] { ".xmind" }.Any(path.ToLower().EndsWith);
        }

        public void Prepare(string path, ContextObject context)
        {
            _imagePath = ExtractFile(path);
            _viewer.Prepare(_imagePath, context);
            WriteFile(logPath, "Prepare successful");
            WriteFile(logPath, "==========================");
        }

        public void View(string path, ContextObject context)
        {
            _imagePath = ExtractFile(path);
            _viewer.View(_imagePath, context);
            WriteFile(logPath, "View successful");
            WriteFile(logPath, context.Title);
            WriteFile(logPath, "==========================");
        }

        public void Cleanup()
        {
            //_ip?.Dispose();
            //_ip = null;
            _viewer.Cleanup();
            _viewer = null;
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