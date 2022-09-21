using System;
using System.Configuration;
using System.IO;

namespace DocumentTagging
{
	class Program
	{
		static void Main(string[] args)
		{
			string sourcedir = ConfigurationManager.AppSettings.Get("sourcedir");
			try
			{
				handleAllFiles(sourcedir, ConfigurationManager.AppSettings.Get("destdir"));
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}

		static void handleAllFiles(string sPath, string dPath)
		{
			DirectoryInfo d = new DirectoryInfo(sPath);
			FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
			if (!Directory.Exists(dPath))
			{
				Directory.CreateDirectory(dPath);
			}

			foreach (FileSystemInfo fsinfo in fsinfos)
			{
				if (fsinfo is DirectoryInfo)
				{
					handleAllFiles(fsinfo.FullName, dPath + "\\" + fsinfo.Name);
				}
				else
				{
					FileInfo fi = new FileInfo(fsinfo.FullName);
					fi.CopyTo(dPath + "\\" + fi.Name, true);
					FileStream fs = new FileStream(dPath + "\\" + fi.Name, FileMode.Open);
					AddinEmbedder.EmbedAddin(getFileType(fsinfo), fs);
					fs.Close();

				}
			}

		}
		static string getFileType(FileSystemInfo file)
		{
			switch (file.Extension.ToLower())
			{
				case ".xlsx":
					return "Excel";
				case ".docx":
					return "Word";
				case ".pptx":
					return "PowerPoint";
				default:
					return null;
			}
		}
	}
}
