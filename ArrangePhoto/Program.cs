using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Shell32;

namespace ArrangePhoto
{
    class Program
    {
        //it isn't stressing the garbage man
        private static Regex r = new Regex(":");
        private static Regex mp4_1 = new Regex(@"^WP_(\d{8})_(\d{2})_(\d{2})_(\d{2})_Pro$");
        private static Regex mp4_2 = new Regex(@"^WP_(\d{8})_(\d{2})(\d{2})(\d{2})Z$");
        private static Regex mp4_3 = new Regex(@"^WP_(\d{8})_\d{3}$");
        private static Regex doneCheck = new Regex(@"^\d{8}_\d{6}$");
        private static Regex epochExpression = new Regex(@"\d{13}");
        public static DateTime GetDateTakenFromMediaInfo(string path)
        {
            Process p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "MediaInfo.exe",
                    Arguments = string.Format("\"{0}\"", path),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };
            p.Start();
            List<string> lines = new List<string>();
            while (!p.StandardOutput.EndOfStream)
            {
                lines.Add(p.StandardOutput.ReadLine());
            }
            //var line = lines.Find(x => x.StartsWith("Recorded date"));
            //if (line != null)
            //{
            //    return DateTime.Parse(line.Substring(line.IndexOf(':') + 1).Trim()).AddHours(-8);
            //}
            var line = lines.Find(x => x.StartsWith("Encoded date"));
            if (line != null)
            {
                string dtStr = line.Substring(line.IndexOf("UTC") + 3).Trim();
                return DateTime.Parse(dtStr).AddHours(8);
            }

            return DateTime.MinValue;
        }
        //retrieves the datetime WITHOUT loading the whole image
        public static DateTime GetDateTakenFromImage(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                using (Image myImage = Image.FromStream(fs, false, false))
                {
                    PropertyItem propItem = myImage.GetPropertyItem(36867);
                    string dateTaken = r.Replace(Encoding.UTF8.GetString(propItem.Value), "-", 2);
                    return DateTime.Parse(dateTaken);
                }
            }
            catch
            {
                return DateTime.MinValue;
            }
        }
        public static DateTime GetDateTakenFromMp4(string path)
        {
            string fileName = Path.GetFileNameWithoutExtension(path);
            Match m = mp4_1.Match(fileName);
            if (m.Success)
            {
                string day = m.Groups[1].Value;
                string hour = m.Groups[2].Value;
                string minute = m.Groups[3].Value;
                string second = m.Groups[4].Value;
                return new DateTime(int.Parse(day.Substring(0, 4)), int.Parse(day.Substring(4, 2)), int.Parse(day.Substring(6)),
                    int.Parse(hour), int.Parse(minute), int.Parse(second));
            }
            m = mp4_2.Match(fileName);
            if (m.Success)
            {
                string day = m.Groups[1].Value;
                string hour = m.Groups[2].Value;
                string minute = m.Groups[3].Value;
                string second = m.Groups[4].Value;
                return new DateTime(int.Parse(day.Substring(0, 4)), int.Parse(day.Substring(4, 2)), int.Parse(day.Substring(6)),
                    int.Parse(hour), int.Parse(minute), int.Parse(second)).AddHours(8);
            }
            return GetDateTakenFromMediaInfo(path);
        }
        public static DateTime GetDateTakenFromAvi(string path)
        {
            Shell32.Shell shell = new Shell32.Shell();
            string folderName = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);
            Folder folder = shell.NameSpace(folderName);
            FolderItem file = folder.ParseName(fileName);

            // These are the characters that are not allowing me to parse into a DateTime
            char[] charactersToRemove = new char[] {
                (char)8206,
                (char)8207
            };

            // Getting the "Media Created" label (don't really need this, but what the heck)
            //string name = folder.GetDetailsOf(null, 191);
            string value = folder.GetDetailsOf(file, 201).Trim();
            //if (string.IsNullOrEmpty(value))
            //{
            //    value = folder.GetDetailsOf(file, 191).Trim();
            //}
            //for (int i = 0; i < 500; ++i)
            //{
            //    value = folder.GetDetailsOf(file, i).Trim();
            //    if (!string.IsNullOrEmpty(value))
            //    {
            //        Console.WriteLine(value);
            //    }
            //}
            // Removing the suspect characters
            foreach (char c in charactersToRemove)
                value = value.Replace((c).ToString(), "").Trim();

            // If the value string is empty, return DateTime.MinValue, otherwise return the "Media Created" date
            return value == string.Empty ? DateTime.MinValue : DateTime.Parse(value);
        }
        public static DateTime GetDateTakenFromEpoch(string path)
        {
            string name = Path.GetFileNameWithoutExtension(path);
            Match m = epochExpression.Match(name);
            if (!m.Success) return DateTime.MinValue;
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            epoch = epoch.AddMilliseconds(Convert.ToDouble(m.Value));
            epoch = epoch.AddHours(8);
            return epoch;
        }
        private static string GenerateFileName(DateTime dt, string file)
        {
            if (dt == DateTime.MinValue)
            {
                return file;
            }
            string extension = Path.GetExtension(file);
            string newFileStem = Path.Combine(Path.GetDirectoryName(file), dt.ToString("yyyyMMdd_HHmmss"));
            string newFile = newFileStem;
            if (Path.GetFileName(file).Equals(Path.GetFileName(newFile) + extension, StringComparison.OrdinalIgnoreCase))
            {
                return newFile + extension;
            }
            int i = 1;
            while (File.Exists(newFile + extension))
            {
                newFile = newFileStem + "_" + i;
                if (Path.GetFileName(file).Equals(Path.GetFileName(newFile) + extension, StringComparison.OrdinalIgnoreCase))
                {
                    return newFile + extension;
                }
                i++;
            }
            newFile = newFile + extension;
            return newFile;
        }
        [STAThread]
        static void Main(string[] args)
        {
            string folder = args[0];
            foreach (var file in Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories))
            {
                if (doneCheck.IsMatch(Path.GetFileNameWithoutExtension(file)))
                {
                    continue;
                }
                DateTime dt = DateTime.MinValue;
                if (Path.GetExtension(file).Equals(".avi", StringComparison.OrdinalIgnoreCase))
                {
                    dt = GetDateTakenFromAvi(file);
                }
                else if (Path.GetExtension(file).Equals(".jpg", StringComparison.OrdinalIgnoreCase)
                    || Path.GetExtension(file).Equals(".jpeg", StringComparison.OrdinalIgnoreCase))
                {
                    dt = GetDateTakenFromImage(file);
                }
                else if (Path.GetExtension(file).Equals(".mp4", StringComparison.OrdinalIgnoreCase))
                {
                    dt = GetDateTakenFromMp4(file);
                }
                else if (Path.GetExtension(file).Equals(".mov", StringComparison.OrdinalIgnoreCase))
                {
                    dt = GetDateTakenFromMediaInfo(file);
                }
                if (dt == DateTime.MinValue)
                {
                    dt = GetDateTakenFromEpoch(file);
                }
                if (dt == DateTime.MinValue)
                {
                    continue;
                }
                string newFile = GenerateFileName(dt, file);
                if (Path.GetFileName(file).Equals(Path.GetFileName(newFile), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                Console.WriteLine("File {0} Date Token {1}", file, dt);
                Console.WriteLine("Rename File {0} to {1}", Path.GetFileName(file), Path.GetFileName(newFile));
                File.Move(file, newFile);
            }
        }


    }
}
