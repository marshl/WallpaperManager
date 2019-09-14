using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
        private DirectoryInfo GetWallpaperFolder()
        {
            var userDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return new DirectoryInfo(Path.Combine(userDir, "Pictures", "CustomWallpapers"));
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
                if (this.CanFileBeAdded(file) || this.CanFileBeRemoved(file))
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

            foreach (string path in this.SelectedItemPaths)
            {
                FileInfo file = new FileInfo(path);
                canAddFile |= this.CanFileBeAdded(file);
                canRemoveFile |= this.CanFileBeRemoved(file);
            }

            //  Create the menu strip
            var menu = new ContextMenuStrip();

            if (canAddFile)
            {
                //  Create a 'count lines' item
                var itemAddWalpaper = new ToolStripMenuItem
                {
                    Text = "Add to wallpapers",
                    //Image = Properties.Resources.CountLines
                };

                //  When we click, we'll count the lines
                itemAddWalpaper.Click += (sender, args) => AddAsWallpaper();
                //  Add the item to the context menu.
                menu.Items.Add(itemAddWalpaper);
            }

            if (canRemoveFile)
            {
                var itemRemoveWallpaper = new ToolStripMenuItem
                {
                    Text = "Remove from wallpapers"
                };

                itemRemoveWallpaper.Click += (sender, args) => RemoveWallpapers();
                menu.Items.Add(itemRemoveWallpaper);
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
                Process converter = new Process();
                converter.StartInfo.FileName = @"C:\Program Files (x86)\ImageMagick-6.9.3-Q16\convert.exe";
                converter.StartInfo.Arguments = $"\"{file.FullName}\" -resize 1920x1080 \"{Path.Combine(wallpaperFolder.FullName, outputFileName)}\"";
                
                var confirmResult = MessageBox.Show($"Run convert.exe {converter.StartInfo.Arguments}",
                                     "Confirm copy",
                                     MessageBoxButtons.YesNo);

                if (confirmResult == DialogResult.Yes)
                {
                    converter.Start();
                }
            }
            else
            {
                var confirmResult = MessageBox.Show($"Copy to {Path.Combine(wallpaperFolder.FullName, file.Name)}",
                                     "Confirm copy",
                                     MessageBoxButtons.YesNo);

                if (confirmResult == DialogResult.Yes)
                {
                    File.Copy(file.FullName, Path.Combine(wallpaperFolder.FullName, file.Name));
                }
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
            var confirmResult = MessageBox.Show($"Are you sure you want to remove {filePathToDelete}?",
                                     "Confirm deletion",
                                     MessageBoxButtons.YesNo);

            if (confirmResult == DialogResult.Yes)
            {
                File.Delete(filePathToDelete);
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
