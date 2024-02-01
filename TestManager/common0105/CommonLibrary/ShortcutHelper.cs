/*
* ShortcutHelper.cs 
* 
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Bencool   <Bencool.lin@quantatw.com>
*/

using System;
using System.IO;
using System.Runtime.InteropServices;


namespace CaptainWin.CommonAPI
{

    /// <summary>
    ///Help to get a shortcut info for a given shortcut path 
    /// </summary>
    public class ShortcutHelper
    {
        [DllImport("shell32.dll")]
        private static extern int SHGetFileInfo(string path, uint fileAttributes, out SHFILEINFO fileInfo, uint cbFileInfo, uint flags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }


        /// <summary>
        /// Giving a shortcut path , return filename, file size info
        /// </summary>
        /// <param name="shortcutPath">short cut file path</param>
        /// <returns>string of file name and file size : $"{fileName}, {fileSize} bytes"</returns>
        public static string GetShortcutInfo(string shortcutPath)
        {
            try
            {
                FileInfo fileInfo = new FileInfo(shortcutPath);

                // Get file name
                string fileName = Path.GetFileName(shortcutPath);

                // Get file size
                long fileSize = fileInfo.Length;

                // Get program name
            //    string programName = GetProgramName(shortcutPath);

                return $"{fileName}, {fileSize} bytes";

                // return $"{fileName}, {fileSize} bytes, {programName}";

            }
            catch (Exception)
            {
                // Handle exceptions, e.g., if the shortcut is invalid
                return null;
            }
        }

        private static string GetProgramName(string shortcutPath)
        {
            SHFILEINFO shinfo = new SHFILEINFO();
            const uint SHGFI_DISPLAYNAME = 0x000000200; // DisplayName
            const uint SHGFI_TYPENAME = 0x000000400; // TypeName

            if (SHGetFileInfo(shortcutPath, 0, out shinfo, (uint)Marshal.SizeOf(shinfo), SHGFI_DISPLAYNAME | SHGFI_TYPENAME) != 0)
            {
                return shinfo.szDisplayName;
            }

            return null;
        }
    }
}
