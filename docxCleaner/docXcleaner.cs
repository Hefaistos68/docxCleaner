using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

using SixLabors.ImageSharp;

namespace docxCleaner
{
	internal class docXcleaner
	{
		private const string KeyLanguage = "-lang";
		private const string KeyStyles = "-styles";
		private const string KeyPrivacy = "-privacy";
		private const string KeyCompress = "-compressimages";
		private const string KeyConvertImages = "-convertimages";
		private const string KeyTitle = "-title";
		private const string KeyRemoveCustom = "-removecustom";
		private const string KeyCompany = "-company";
		private const string KeyCreator = "-creator";
		
		private static Dictionary<string, string> commandlineArgs = new Dictionary<string, string>();

		static void Main(string[] args)
		{
			PrepareArguments(args);

			if (args.Length == 0)
			{
				Console.WriteLine("Please provide a filename as an argument.");
				return;
			}

			string fileName = commandlineArgs["-file"];

			if (!File.Exists(fileName))
			{
				Console.WriteLine($"File '{fileName}' does not exist.");
				return;
			}


			try
			{
				// make a copy of the file before processing
				string backupFileName = Path.ChangeExtension(fileName, ".updated.zip");
				File.Copy(fileName, backupFileName, true);

				using (ZipArchive archive = ZipFile.Open(backupFileName, ZipArchiveMode.Update))
				{
					Console.WriteLine($"Successfully opened '{backupFileName}' as a Zip archive.");

					// update settings file
					if(commandlineArgs.TryGetValue(KeyPrivacy, out string? settingsFile) && !string.IsNullOrEmpty(settingsFile))
					{
						ReplaceSettings(archive, settingsFile);
					}
					else
					{
						ProcessSettings(archive);
					}

					ProcessApp(archive);
					ProcessCore(archive);

					// read the "rels\.rels" file as XML document
					ProcessDocumentRelationships(archive);
					// ProcessRelationships(archive);

					// get all files in the "customXml" folder
					if (commandlineArgs.ContainsKey(KeyRemoveCustom))
					{
						ProcessCustomXmlFolder(archive);
					}

					if (commandlineArgs.ContainsKey(KeyStyles))
					{
						ProcessStyles(archive);
					}

					// find a file called "custom.xml" in the "docProps" folder
					// ProcessCustomXmlFile(archive);

					// read the "[Content_Types].xml" file as XML document
					ProcessContentTypes(archive);

					// update all paragraphs with correct language setting
					if(commandlineArgs.TryGetValue(KeyLanguage, out string? lang) && !string.IsNullOrEmpty(lang))
					{
						ProcessLanguageSettingDocument(archive, lang);
						ProcessLanguageSettingHeader(archive, lang);
						ProcessLanguageSettingFooter(archive, lang);
					}

					ProcessStyles(archive);

					// compress all images
					if(commandlineArgs.ContainsKey(KeyConvertImages))
					{
						if (int.TryParse(commandlineArgs[KeyCompress], out int quality))
						{
							CompressImages(archive, quality);
						}
						else
						{
							Console.WriteLine("Invalid or missing JPG quality value. Using default value of 75.");
							CompressImages(archive, 75);
						}
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"An error occurred while opening the file: {ex.Message}");
			}
		}

		/// <summary>
		/// Replaces the settings file in the given Zip archive with the specified settings file.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="settingsFile">The path to the settings file.</param>
		private static void ReplaceSettings(ZipArchive archive, string settingsFile)
		{
			string fileName = "word/settings.xml";

			DeleteFile(archive, fileName);
			AddFile(archive, fileName, File.ReadAllBytes(settingsFile));
		}

		/// <summary>
		/// Prepares the command line arguments.
		/// </summary>
		/// <param name="args">The command line arguments.</param>
		private static void PrepareArguments(string[] args)
		{
			int n = 0;

			if (args.Length >= 1)
			{
				if (args[0] == "?" || args[0] == "-help" || args[0] == "-?")
				{
					Console.WriteLine("Usage: docxCleaner -file <filename> [-lang <language>] [-styles <filename>] [-custom <filename>] [-compressimages <quality>]");
					Console.WriteLine("  -file <filename>           The name of the file to process");
					Console.WriteLine("  -lang <language>           The language to set for the document");
					Console.WriteLine("  -styles <filename>         The name of the file containing the styles to apply");
					Console.WriteLine("  -custom <filename>         The name of the file containing the custom settings");
					Console.WriteLine("  -compressimages <quality>  The quality to use when compressing images");
					Console.WriteLine("  -privacy                   Remove personal information");
					Console.WriteLine("  -convertimages             Convert all images to JPG format");
					Console.WriteLine("  -removecustom              Remove all customXml content");
					Console.WriteLine("  -title                     Set a document title property (or remove if empty)");
					Console.WriteLine("  -company                   Set a company property (or remove if empty)");
					Console.WriteLine("  -creator                   Set a creator property (or remove if empty)");

					return;
				}

				commandlineArgs.Add("-file", args[n++]);
			}

			while (n < args.Length)
			{
				string key = args[n++];
				string value = string.Empty;

				if (n < args.Length && !args[n].StartsWith('-'))
				{
					value = args[n++];
				}

				commandlineArgs.Add(key, value);
			}

			// setup default values
			commandlineArgs.TryAdd(KeyCompress, "75");	 // this is the JPG compression value

		}

		private static void ProcessStyles(ZipArchive archive)
		{
			throw new NotImplementedException();
		}

		/// <summary>
		/// Processes the settings file in the given Zip archive.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		private static void ProcessSettings(ZipArchive archive)
		{
			string fileName = "word/settings.xml";

			ZipArchiveEntry? wordDocumentEntry = archive.GetEntry(fileName);
			if (wordDocumentEntry != null)
			{
				Console.WriteLine($"Found '{fileName}' file.");

				// read file wordDocumentEntry as Xml document
				bool changed = false;
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

				using (Stream stream = wordDocumentEntry.Open())
				{
					doc.Load(stream);

					// find all nodes "w:lang" with attribute "w:val" and set "w:val" to "en-GB"
					XmlNode? node = doc.SelectSingleNode("//w:settings", nsmgr);
					if (node != null)
					{
						// add a new child node "w:removePersonalInformation" 
						XmlElement newNode = doc.CreateElement("w:removePersonalInformation", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
						node.AppendChild(newNode);

						newNode = doc.CreateElement("w:removeDateAndTime", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
						node.AppendChild(newNode);

						changed = true;
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, fileName, doc);
				}
			}
		}

		/// <summary>
		/// Processes the app.xml file in the given Zip archive, sets the title and company properties.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		private static void ProcessApp(ZipArchive archive)
		{
			string fileName = "docProps/app.xml";

			ZipArchiveEntry? wordDocumentEntry = archive.GetEntry(fileName);
			if (wordDocumentEntry != null)
			{
				Console.WriteLine($"Found '{fileName}' file.");

				// read file wordDocumentEntry as Xml document
				bool changed = false;
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("ep", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
				nsmgr.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

				using (Stream stream = wordDocumentEntry.Open())
				{
					doc.Load(stream);

					if (commandlineArgs.TryGetValue(KeyCompany, out string? company))
					{
						XmlNode? node = doc.SelectSingleNode("//ep:Company", nsmgr);
						if (node != null)
						{
							node.InnerText = company;

							changed = true;
						}
					}

					if (commandlineArgs.TryGetValue(KeyTitle, out string? title))
					{
						XmlNode? node = doc.SelectSingleNode("//ep:TitlesOfParts/vt:vector/vt:lpstr", nsmgr);
						if (node != null)
						{
							node.InnerText = title;

							changed = true;
						}
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, fileName, doc);
				}
			}
		}

		/// <summary>
		/// Processes the core.xml file in the given Zip archive. Replaces the title and creator properties and clears the last modified, last printed, created and modified properties.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="title">The new title for the document.</param>
		private static void ProcessCore(ZipArchive archive)
		{
			string fileName = "docProps/core.xml";

			ZipArchiveEntry? wordDocumentEntry = archive.GetEntry(fileName);
			if (wordDocumentEntry != null)
			{
				Console.WriteLine($"Found '{fileName}' file.");

				// read file wordDocumentEntry as Xml document
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
				nsmgr.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
				nsmgr.AddNamespace("dcterms", "http://purl.org/dc/terms/");

				using (Stream stream = wordDocumentEntry.Open())
				{
					doc.Load(stream);

					XmlNode? node;

					if(commandlineArgs.TryGetValue(KeyTitle, out string? title))
					{
						node = doc.SelectSingleNode("//dc:title", nsmgr);
						if (node != null)
							node.InnerText = title;
					}

					if (commandlineArgs.TryGetValue(KeyCreator, out string? creator))
					{
						node = doc.SelectSingleNode("//dc:creator", nsmgr);
						if (node != null)
							node.InnerText = creator;
					}

					if(commandlineArgs.ContainsKey(KeyPrivacy))
					{
						node = doc.SelectSingleNode("//cp:lastModifiedBy", nsmgr);
						if (node != null)
							node.InnerText = string.Empty;
						node = doc.SelectSingleNode("//cp:lastPrinted", nsmgr);
						if (node != null)
							node.InnerText = string.Empty;
						node = doc.SelectSingleNode("//dcterms:created", nsmgr);
						if (node != null)
							node.InnerText = string.Empty;
						node = doc.SelectSingleNode("//dcterms:modified", nsmgr);
						if (node != null)
							node.InnerText = string.Empty;
					}
				}

				ReplaceFileContent(archive, fileName, doc);
			}
		}

		/// <summary>
		/// Compresses the images in the specified Zip archive with the given image quality.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="imageQuality">The quality to use when compressing images.</param>
		private static void CompressImages(ZipArchive archive, int imageQuality)
		{
			// get all files in the "word/media" folder
			List<ZipArchiveEntry> imageEntries = archive.Entries.Where(e => e.FullName.StartsWith("word/media/")).ToList();

			// load each image file and compress it
			foreach (ZipArchiveEntry entry in imageEntries)
			{
				Console.WriteLine($"Compressing image '{entry.FullName}'.");

				MemoryStream newImageStream = new MemoryStream();

				using (Stream stream = entry.Open())
				{

					using (MemoryStream ms = new MemoryStream())
					{
						stream.CopyTo(ms);
						ms.Seek(0, SeekOrigin.Begin);
						using (Image image = Image.Load(ms))
						{
							// save the image back to the stream with a lower quality
							image.Save(newImageStream, new SixLabors.ImageSharp.Formats.Jpeg.JpegEncoder()
							{
								Quality = imageQuality
							});
						}
					}
				}

				// replace the image in the archive with the compressed image
				string newFilename = Path.ChangeExtension(entry.FullName, ".jpg");

				AddFile(archive, newFilename, newImageStream.ToArray());
				DeleteFile(archive, entry.FullName);

				string fileNameOld = Path.GetFileName(entry.FullName);
				string fileNameNew = Path.GetFileName(newFilename);

				UpdateReferences(archive, fileNameOld, fileNameNew);
			}
		}

		/// <summary>
		/// Updates the references to the specified image file in the given Zip archive.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="fullName">The full name of the image file to be replaced.</param>
		/// <param name="newFilename">The new filename for the image file.</param>
		private static void UpdateReferences(ZipArchive archive, string fullName, string newFilename)
		{
			// we need to update document.xml, header1.xml and footer1.xml
			ProcessImageReference(archive, "word/_rels/document.xml.rels", fullName, newFilename);
			ProcessImageReference(archive, "word/_rels/header1.xml.rels", fullName, newFilename);
			ProcessImageReference(archive, "word/_rels/footer1.xml.rels", fullName, newFilename);
		}

		/// <summary>
		/// Processes the image reference in the given Zip archive.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="fileName">The name of the file containing the image references.</param>
		/// <param name="fullName">The full name of the image file to be replaced.</param>
		/// <param name="newFilename">The new filename for the image file.</param>
		private static void ProcessImageReference(ZipArchive archive, string fileName, string fullName, string newFilename)
		{
			ZipArchiveEntry? relsEntry = archive.GetEntry(fileName);
			if (relsEntry != null)
			{
				Console.WriteLine($"Found '{fileName}' file.");

				// read file wordDocumentEntry as Xml document
				bool changed = false;
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

				using (Stream stream = relsEntry.Open())
				{
					doc.Load(stream);

					// find all "Relationship" nodes
					XmlNodeList? nodes = doc.SelectNodes("//r:Relationship[contains(@Target, 'media/')]", nsmgr);

					// delete all nodes and write XML document back to archive
					if (nodes != null)
					{
						foreach (XmlNode node in nodes)
						{
							Console.WriteLine($"Found an image reference node with Target='{node.Attributes["Target"].Value}'.");
							node.Attributes["Target"].Value = node.Attributes["Target"].Value.Replace(fullName, newFilename);
						}

						changed = nodes.Count >= 1;
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, fileName, doc);
				}
			}
		}

		/// <summary>
		/// Processes the language setting in the given Zip archive.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="targetLanguage">The target language to set.</param>
		/// <param name="fileName">The name of the file to process.</param>
		private static void ProcessLanguageSetting(ZipArchive archive, string targetLanguage, string fileName)
		{
			ZipArchiveEntry? wordDocumentEntry = archive.GetEntry(fileName);
			if (wordDocumentEntry != null)
			{
				Console.WriteLine($"Found '{fileName}' file.");

				// read file wordDocumentEntry as Xml document
				bool changed = false;
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

				using (Stream stream = wordDocumentEntry.Open())
				{
					doc.Load(stream);

					// find all nodes "w:lang" with attribute "w:val" and set "w:val" to the target language
					XmlNodeList? nodes = doc.SelectNodes("//w:lang", nsmgr);
					if (nodes != null)
					{
						foreach (XmlNode node in nodes)
						{
							Console.WriteLine($"Found a lang node with value='{node.Attributes["w:val"].Value}'.");
							node.Attributes["w:val"].Value = targetLanguage;
						}

						changed = true;
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, fileName, doc);
				}
			}
		}
		
		private static void ProcessLanguageSettingDocument(ZipArchive archive, string targetLanguage)
		{
			ProcessLanguageSetting(archive, targetLanguage, "word/document.xml");
		}

		/// <summary>
		/// Processes the language setting in the given Zip archive for all footer files.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="targetLanguage">The target language to set.</param>
		private static void ProcessLanguageSettingFooter(ZipArchive archive, string targetLanguage)
		{
			List<ZipArchiveEntry> footerEntries = archive.Entries.Where(e => e.FullName.StartsWith("word/footer") && e.FullName.EndsWith(".xml")).ToList();

			foreach (ZipArchiveEntry entry in footerEntries)
			{
				ProcessLanguageSetting(archive, targetLanguage, entry.FullName);
			}
			// ProcessLanguageSetting(archive, targetLanguage, "word/footer1.xml");
		}

		/// <summary>
		/// Processes the language setting in the given Zip archive for all header files.
		/// </summary>
		/// <param name="archive">The Zip archive.</param>
		/// <param name="targetLanguage">The target language to set.</param>
		private static void ProcessLanguageSettingHeader(ZipArchive archive, string targetLanguage)
		{
			List<ZipArchiveEntry> footerEntries = archive.Entries.Where(e => e.FullName.StartsWith("word/header") && e.FullName.EndsWith(".xml")).ToList();

			foreach (ZipArchiveEntry entry in footerEntries)
			{
				ProcessLanguageSetting(archive, targetLanguage, entry.FullName);
			}
			// ProcessLanguageSetting(archive, targetLanguage, "word/header1.xml");
		}

		private static void ProcessContentTypes(ZipArchive archive)
		{
			string fileName = "[Content_Types].xml";

			ZipArchiveEntry? contentTypes = archive.GetEntry(fileName);
			if (contentTypes != null)
			{
				Console.WriteLine($"Found '{fileName}' file.");

				bool changed = false;

				// read file wordDocumentEntry as Xml document
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/content-types");

				using (Stream stream = contentTypes.Open())
				{
					doc.Load(stream);

					// find all nodes "Override" with attribute "PartName" contains "/customXml/"
					XmlNodeList nodes = doc.SelectNodes("//ns:Override[contains(@PartName, '/customXml/')]", nsmgr);

					// delete all nodes referencing customXML
					if (nodes != null)
					{
						foreach (XmlNode node in nodes)
						{
							Console.WriteLine($"Found an Override node with PartName='{node.Attributes["PartName"].Value}'.");
							node.ParentNode.RemoveChild(node);
						}

						changed = true;
					}

					// find a Default node for JPG extension, or insert a new one
					if (commandlineArgs.ContainsKey(KeyCompress))
					{
						XmlNode? defaultNode = doc.SelectSingleNode("//ns:Default[@Extension='jpg']", nsmgr);
						if (defaultNode == null)
						{
							// create a new Default content type node for JPG 
							XmlElement newNode = doc.CreateElement("Default", "http://schemas.openxmlformats.org/package/2006/content-types");
							newNode.SetAttribute("Extension", "jpg");
							newNode.SetAttribute("ContentType", "image/jpeg");

							// append this node to the "Types" node
							XmlNode typesNode = doc.SelectSingleNode("//ns:Types", nsmgr);
							typesNode!.AppendChild(newNode);

							changed = true;
						}
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, "[Content_Types].xml", doc);
				}
			}
		}

		private static void ProcessCustomXmlFile(ZipArchive archive)
		{
			ZipArchiveEntry? customXmlEntry = archive.GetEntry("docProps/custom.xml");
			if (customXmlEntry != null)
			{
				Console.WriteLine("Found 'custom.xml' in the 'docProps' folder.");
				customXmlEntry.Delete();
			}
		}

		private static void ProcessCustomXmlFolder(ZipArchive archive)
		{
			List<ZipArchiveEntry> customXmlEntries = archive.Entries.Where(e => e.FullName.StartsWith("customXml/")).ToList();

			// delete all files in the customXmlEntries list
			foreach (ZipArchiveEntry entry in customXmlEntries)
			{
				Console.WriteLine($"Deleting file '{entry.FullName}' from the archive.");
				entry.Delete();
			}
		}

		private static void ProcessRelationships(ZipArchive archive)
		{
			ZipArchiveEntry? relsEntry = archive.GetEntry("_rels/.rels");
			if (relsEntry != null)
			{
				Console.WriteLine("Found '_rels/.rels' file.");

				// read file wordDocumentEntry as Xml document
				bool changed = false;
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

				using (Stream stream = relsEntry.Open())
				{
					doc.Load(stream);

					// find all "Relationship" nodes
					XmlNodeList nodes = doc.SelectNodes("//r:Relationship[contains(@Target, '/custom.xml')]", nsmgr);

					// delete all nodes and write XML document back to archive
					if (nodes != null)
					{
						foreach (XmlNode node in nodes)
						{
							Console.WriteLine($"Found an Override node with Target='{node.Attributes["Target"].Value}'.");
							node.ParentNode.RemoveChild(node);
						}

						changed = true;
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, "_rels/.rels", doc);
				}
			}
		}

		private static void ProcessDocumentRelationships(ZipArchive archive)
		{
			ZipArchiveEntry? relsEntry = archive.GetEntry("word/_rels/document.xml.rels");
			if (relsEntry != null)
			{
				Console.WriteLine("Found 'word/_rels/document.xml.rels' file.");

				// read file wordDocumentEntry as Xml document
				bool changed = false;
				XmlDocument doc = new XmlDocument();
				XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
				nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

				using (Stream stream = relsEntry.Open())
				{
					doc.Load(stream);

					// find all "Relationship" nodes
					XmlNodeList nodes = doc.SelectNodes("//r:Relationship[contains(@Target, '/customXml/')]", nsmgr);

					// delete all nodes and write XML document back to archive
					if (nodes != null)
					{
						foreach (XmlNode node in nodes)
						{
							Console.WriteLine($"Found an Override node with Target='{node.Attributes["Target"].Value}'.");
							node.ParentNode.RemoveChild(node);
						}

						changed = true;
					}
				}

				if (changed)
				{
					ReplaceFileContent(archive, "word/_rels/document.xml.rels", doc);
				}
			}
		}

		private static void ReplaceFileContent(ZipArchive archive, string entryName, XmlDocument newContent)
		{
			ReplaceFileContent(archive, entryName, Encoding.UTF8.GetBytes(newContent.OuterXml));
		}

		private static void ReplaceFileContent(ZipArchive archive, string entryName, byte[] newContent)
		{
			// Find the entry to replace
			ZipArchiveEntry? entry = archive.GetEntry(entryName);
			if (entry != null)
			{
				// Delete the existing entry
				entry.Delete();
			}

			// Create a new entry with the same name
			ZipArchiveEntry newEntry = archive.CreateEntry(entryName);

			// Write the new content to the new entry
			using (Stream stream = newEntry.Open())
			{
				stream.Write(newContent, 0, newContent.Length);
			}
		}

		private static void DeleteFile(ZipArchive archive, string entryName)
		{
			// Find the entry to delete
			ZipArchiveEntry? entry = archive.GetEntry(entryName);
			if (entry != null)
			{
				// Delete the existing entry
				entry.Delete();
			}
		}

		private static void AddFile(ZipArchive archive, string entryName, byte[] content)
		{
			// Create a new entry with the specified name
			ZipArchiveEntry newEntry = archive.CreateEntry(entryName);
			// Write the content to the new entry
			using (Stream stream = newEntry.Open())
			{
				stream.Write(content, 0, content.Length);
			}
		}
	}
}
