using ImageMagick;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;



namespace WallpaperManagementExtension
{

    /// <summary>
    /// The CountLinesExtensions is an example shell context menu extension,
    /// implemented with SharpShell. It adds the command 'Count Lines' to text
    /// files.
    /// </summary>
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    public class WallpaperManagementExtension : SharpContextMenu
    {
        private DirectoryInfo GetPictureFolder()
        {
            var userDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return new DirectoryInfo(Path.Combine(userDir, "Pictures"));
        }

        private DirectoryInfo GetWallpaperFolder()
        {
            return new DirectoryInfo(Path.Combine(this.GetPictureFolder().FullName, "CustomWallpapers"));
        }

        private bool IsInPictures(FileInfo file)
        {
            return this.IsFileInDirectory(file, "Pictures");
        }

        private bool IsInCustomWallpapers(FileInfo file)
        {
            return this.IsFileInDirectory(file, "CustomWallpapers");
        }

        private bool IsFileInDirectory(FileInfo file, string directoryName)
        {
            if (!file.Exists)
            {
                return false;
            }

            for (DirectoryInfo parentDir = file.Directory; parentDir != null; parentDir = parentDir.Parent)
            {
                if (parentDir.Name == directoryName)
                {
                    return true;
                }
            }
            return false;
        }

        private bool CanFileBeAdded(FileInfo file)
        {
            if (!this.IsInPictures(file) || this.IsInCustomWallpapers(file))
            {
                return false;
            }

            if (this.AcceptedWallpaperFilenames.Contains(Path.GetFileNameWithoutExtension(file.Name)))
            {
                return false;
            }
            return true;
        }

        private bool CanFileBeRemoved(FileInfo file)
        {
            if (!this.IsInPictures(file) || this.IsInCustomWallpapers(file))
            {
                return false;
            }

            if (this.AcceptedWallpaperFilenames.Contains(Path.GetFileNameWithoutExtension(file.Name)))
            {
                return true;
            }
            return false;
        }


        private bool CanFileGotoSource(FileInfo fileInfo)
        {
            return this.IsInCustomWallpapers(fileInfo);
        }

        private readonly List<string> AcceptedWallpaperFilenames = new List<string>();

        private void GetAcceptedWallpaperFiles()
        {
            var wallpaperFiles = this.GetWallpaperFolder().GetFiles();
            foreach (var wallpaper in wallpaperFiles)
            {
                this.AcceptedWallpaperFilenames.Add(Path.GetFileNameWithoutExtension(wallpaper.Name));
            }
        }

        /// <summary>
        /// Determines whether this instance can a shell
        /// context show menu, given the specified selected file list
        /// </summary>
        /// <returns>
        /// <c>true</c> if this instance should show a shell context
        /// menu for the specified file list; otherwise, <c>false</c>
        /// </returns>
        protected override bool CanShowMenu()
        {
            this.GetAcceptedWallpaperFiles();

            foreach (string path in this.SelectedItemPaths)
            {
                FileInfo file = new FileInfo(path);
                if (this.CanFileBeAdded(file) || this.CanFileBeRemoved(file) || this.CanFileGotoSource(file))
                {
                    return true;
                }
            }

            return false;
        }


        /// <summary>
        /// Creates the context menu. This can be a single menu item or a tree of them.
        /// </summary>
        /// <returns>
        /// The context menu for the shell context menu.
        /// </returns>
        protected override ContextMenuStrip CreateMenu()
        {
            this.GetAcceptedWallpaperFiles();
            bool canAddFile = false;
            bool canRemoveFile = false;
            bool canGotoSourceFile = false;

            foreach (string path in this.SelectedItemPaths)
            {
                FileInfo file = new FileInfo(path);
                canAddFile |= this.CanFileBeAdded(file);
                canRemoveFile |= this.CanFileBeRemoved(file);
                canGotoSourceFile |= this.CanFileGotoSource(file);
            }

            //  Create the menu strip
            var menu = new ContextMenuStrip();

            if (canAddFile)
            {
                //  Create a 'count lines' item
                var itemAddWalpaper = new ToolStripMenuItem
                {
                    Text = "Add to wallpapers",
                    Image = Properties.Resources.picture_add,
                };

                itemAddWalpaper.Click += (sender, args) => AddAsWallpaper();
                menu.Items.Add(itemAddWalpaper);
            }

            if (canRemoveFile)
            {
                var itemRemoveWallpaper = new ToolStripMenuItem
                {
                    Text = "Remove from wallpapers",
                    Image = Properties.Resources.picture_remove
                };

                itemRemoveWallpaper.Click += (sender, args) => RemoveWallpapers();
                menu.Items.Add(itemRemoveWallpaper);
            }

            if (canGotoSourceFile)
            {
                var gotoSourceMenuItem = new ToolStripMenuItem
                {
                    Text = "Go to source wallpaper",
                    Image = Properties.Resources.picture_goto,
                };

                gotoSourceMenuItem.Click += (sender, args) => GotoAllWallpapers();
                menu.Items.Add(gotoSourceMenuItem);
            }

            //  Return the menu
            return menu;
        }

        /// <summary>
        /// Counts the lines in the selected files
        /// </summary>
        private void AddAsWallpaper()
        {
            var wallpaperFolder = this.GetWallpaperFolder();
            //  Go through each file.
            foreach (var filePath in SelectedItemPaths)
            {
                FileInfo file = new FileInfo(filePath);
                if (this.CanFileBeAdded(file))
                {
                    this.MoveFileToWallpapers(file, wallpaperFolder);
                }
            }
            this.RefreshOverlays();
        }

        private void MoveFileToWallpapers(FileInfo file, DirectoryInfo wallpaperFolder)
        {
            if (file.Extension != ".jpg")
            {
                string outputFileName = Path.GetFileNameWithoutExtension(file.Name) + ".jpg";
                string outputPath = Path.Combine(wallpaperFolder.FullName, outputFileName);
                this.ConvertImage(file, new FileInfo(outputPath));
            }
            else
            {
                File.Copy(file.FullName, Path.Combine(wallpaperFolder.FullName, file.Name));
            }
        }

        private void ConvertImage(FileInfo source, FileInfo target)
        {
            MagickReadSettings settings = new MagickReadSettings
            {
                Width = 1920,
                Height = 1080
            };
            using (MagickImage image = new MagickImage(source.FullName, settings))
            {
                image.Write(target.FullName);
            }
        }

        private void RemoveWallpapers()
        {
            var wallpaperFolder = this.GetWallpaperFolder();
            //  Go through each file.
            foreach (var filePath in SelectedItemPaths)
            {
                FileInfo file = new FileInfo(filePath);
                if (this.CanFileBeRemoved(file))
                {
                    this.RemoveSingleWWallpaper(file, wallpaperFolder);
                }
            }
            this.RefreshOverlays();
        }

        private void RemoveSingleWWallpaper(FileInfo file, DirectoryInfo wallpaperFolder)
        {
            string filePathToDelete = Path.Combine(wallpaperFolder.FullName, Path.GetFileNameWithoutExtension(file.Name) + ".jpg");
            File.Delete(filePathToDelete);
        }

        private void GotoAllWallpapers()
        {
            DirectoryInfo picturesFolder = this.GetPictureFolder();
            var allFilepaths = picturesFolder.GetFiles("*", SearchOption.AllDirectories);
            foreach (var selectedPath in this.SelectedItemPaths)
            {
                string selectedFilename = Path.GetFileNameWithoutExtension(selectedPath);
                var matchingFiles = allFilepaths.Where(x => Path.GetFileNameWithoutExtension(x.FullName) == selectedFilename && x.FullName != selectedPath);
                foreach (var matchingFile in matchingFiles)
                {
                    Process explorer = new Process();
                    explorer.StartInfo.FileName = "explorer";
                    explorer.StartInfo.Arguments = $"/select,{matchingFile.FullName}";
                    explorer.Start();
                }
            }
        }

        private void RefreshOverlays()
        {
            Process ie4uinit = new Process();
            ie4uinit.StartInfo.FileName = "ie4uinit.exe";
            ie4uinit.StartInfo.Arguments = "-show";
            ie4uinit.Start();
        }
    }
}
