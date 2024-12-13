using System;
using System.IO;
using System.IO.Compression;
using System.IO.Hashing;
using System.Text;
using System.Xml;

namespace MsixPacker
{
    /// <summary>
    /// This exception is thrown when MSIX packaging operations fail.
    /// </summary>
    public class MsixPackingException : Exception
    {
        public MsixPackingException(string message) : base(message) { }
        public MsixPackingException(string message, Exception inner) : base(message, inner) { }
    }

    /// <summary>
    /// MSIX Package Creator
    /// </summary>
    public sealed class MsixPacker : IDisposable
    {
        /// <summary>
        /// Initializes an MSIX packer with output file name.
        /// </summary>
        /// <param name="msixFileName">Path of the output MSIX file.</param>
        public MsixPacker(string msixFileName, String basePath)
        {
            m_fs = File.Open(msixFileName, FileMode.Create, FileAccess.ReadWrite);
            m_basePath = basePath;
        }

        /// <summary>
        /// Base directory path, to which archived file names are make relative.
        /// </summary>
        public string BasePath
        {
            get { return m_basePath; }
            set
            {
                m_basePath = Path.GetFullPath(value);
            }
        }

        /// <summary>
        /// Add a file/directory to the archive.
        /// </summary>
        /// <param name="fileName">Path to a file/directory to archive.</param>
        /// <exception cref="ArgumentException">
        /// The file/directory is outside the <see cref="ZipArc.BasePath"/>.
        /// </exception>
        public void AddFile(string fileName)
        {
            // Skip signature-related files
            string baseName = Path.GetFileName(fileName).ToLowerInvariant();
            if (baseName == "appxmetadata" || baseName == "appxsignature.p7x" || baseName == "codeintegrity.cat" || baseName == "[Content_Types].xml" || baseName == "AppxBlockMap.xml")
                return;

            if (!AddFileInternal(fileName, GenerateInternalFileName(fileName)))
                return;

            if (Directory.Exists(fileName))
            {
                foreach (string fn in Directory.GetDirectories(fileName, "*", SearchOption.AllDirectories))
                    AddFile(fn);

                foreach (string fn in Directory.GetFiles(fileName, "*", SearchOption.AllDirectories))
                    AddFile(fn);
            }
        }

        /// <summary>
        /// Finalizes archiving.
        /// </summary>
        /// <summary>
        /// Generates the AppxBlockMap.xml content for MSIX packaging
        /// </summary>
        public void WriteBlockMap(Stream output)
        {
            var ms = new MemoryStream();
            using (var writer = new XmlTextWriter(ms, Encoding.UTF8))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                writer.WriteStartElement("BlockMap");
                writer.WriteAttributeString("xmlns", "http://schemas.microsoft.com/appx/2010/blockmap");
                writer.WriteAttributeString("HashMethod", "SHA256");

                foreach (var entry in m_entries)
                {
                    if (!entry.IsDirectory)
                    {
                        writer.WriteStartElement("File");
                        writer.WriteAttributeString("Name", entry.Name);
                        writer.WriteAttributeString("Size", entry.UncompressedSize.ToString());
                        writer.WriteAttributeString("LfhSize", entry.LfhSize.ToString());
                        writer.WriteEndElement();
                    }
                }

                writer.WriteEndElement(); // BlockMap
                writer.WriteEndDocument();
                writer.Flush();

                // Copy to output stream
                ms.Position = 0;
                ms.CopyTo(output);
            }
        }

        /// <summary>
        /// Writes the content types XML required for MSIX packages
        /// </summary>
        private void WriteContentTypes(Stream output)
        {
            var ms = new MemoryStream();
            using (var writer = new XmlTextWriter(ms, Encoding.UTF8))
            {
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                writer.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");

                // Default content type for XML files
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "xml");
                writer.WriteAttributeString("ContentType", "application/xml");
                writer.WriteEndElement();

                // Default content type for DLL files
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "dll");
                writer.WriteAttributeString("ContentType", "application/x-msdownload");
                writer.WriteEndElement();

                // Default content type for EXE files
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "exe");
                writer.WriteAttributeString("ContentType", "application/x-msdownload");
                writer.WriteEndElement();

                // Add override for AppxBlockMap
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/AppxBlockMap.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.ms-appx.blockmap+xml");
                writer.WriteEndElement();

                // Add override for AppxManifest
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/AppxManifest.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.ms-appx.manifest+xml");
                writer.WriteEndElement();

                // Add content types for all other files based on extension
                var processedExtensions = new HashSet<string> { "xml", "dll", "exe" };
                foreach (var entry in m_entries)
                {
                    if (!entry.IsDirectory)
                    {
                        string fullExt = Path.GetExtension(entry.Name).TrimStart('.');
                        string[] extParts = fullExt.Split('.');

                        // Handle resource variants (e.g., .scale-100.png)
                        string baseExt = extParts[extParts.Length - 1];
                        if (!processedExtensions.Contains(baseExt))
                        {
                            processedExtensions.Add(baseExt);
                            writer.WriteStartElement("Default");
                            writer.WriteAttributeString("Extension", baseExt);
                            writer.WriteAttributeString("ContentType", GetContentType(baseExt));
                            writer.WriteEndElement();
                        }

                        // Also register the full extension if it's a resource variant
                        if (extParts.Length > 1)
                        {
                            string resourceExt = string.Join(".", extParts);
                            if (!processedExtensions.Contains(resourceExt))
                            {
                                processedExtensions.Add(resourceExt);
                                writer.WriteStartElement("Default");
                                writer.WriteAttributeString("Extension", resourceExt);
                                writer.WriteAttributeString("ContentType", GetContentType(baseExt));
                                writer.WriteEndElement();
                            }
                        }
                    }
                }

                writer.WriteEndElement(); // Types
                writer.WriteEndDocument();
                writer.Flush();

                // Copy to output stream
                ms.Position = 0;
                ms.CopyTo(output);
            }
        }

        /// <summary>
        /// Gets the MIME content type for a file extension using Windows Registry
        /// </summary>
        private string GetContentType(string extension)
        {
            // Handle MSIX-specific content types first
            switch (extension.ToLower())
            {
                // Core MSIX files
                case "blockmap":
                    return "application/vnd.ms-appx.blockmap+xml";
                case "p7x":
                    return "application/vnd.ms-appx.signature";
                case "appxmanifest":
                case "manifest":
                    return "application/vnd.ms-appx.manifest+xml";
                case "appx":
                case "msix":
                case "appxbundle":
                case "msixbundle":
                    return "application/vnd.ms-appx";
            }

            return MimeTypes.GetMimeType(extension);
        }

        public void Dispose()
        {
            WriteCentralDirectory();
            m_fs.Dispose();
            GC.SuppressFinalize(this);
        }

        private bool AddFileInternal(string fileName, string internalName)
        {
            if (string.IsNullOrEmpty(internalName))
                return true;

            // duplication check
            if (m_fileSet.Contains(internalName))
                return false;

            var fe = new FileEntry(fileName, internalName);
            fe.WriteLocalFileHeader(m_fs);
            m_entries.Add(fe);
            m_fileSet.Add(internalName);
            return true;
        }

        private string GenerateInternalFileName(string fileName)
        {
            var internalName = Path.GetFullPath(fileName);
            if (internalName == m_basePath)
                return string.Empty;

            if (!internalName.StartsWith(m_basePath))
                throw new ArgumentException("The file is not in the base directory!", "fileName");

            // make the file name relative to the base path
            internalName = internalName.Substring(m_basePath.Length + 1).Replace(Path.DirectorySeparatorChar, '/');

            // confirms all the ascending directories exist
            var comps = internalName.Split('/');
            var dir = string.Empty;
            for (int i = 0; i < comps.Length - 1; i++)
            {
                dir += comps[i] + "/";
                if (!m_fileSet.Contains(dir))
                    AddFileInternal(Path.GetFullPath(m_basePath + "/" + dir), dir);
            }

            // directory should suffixed by /
            if (Directory.Exists(fileName))
                internalName += "/";

            return internalName;
        }

        private void WriteCentralDirectory()
        {
            // Generate and add block map before central directory
            using (var ms = new MemoryStream())
            {
                WriteBlockMap(ms);
                ms.Position = 0;
                var blockMapEntry = new FileEntry("AppxBlockMap.xml", "AppxBlockMap.xml", ms.ToArray());
                blockMapEntry.WriteLocalFileHeader(m_fs);
                m_entries.Add(blockMapEntry);
            }

            // Generate and add content types
            using (var ms = new MemoryStream())
            {
                WriteContentTypes(ms);
                ms.Position = 0;
                var contentTypesEntry = new FileEntry("[Content_Types].xml", "[Content_Types].xml", ms.ToArray());
                contentTypesEntry.WriteLocalFileHeader(m_fs);
                m_entries.Add(contentTypesEntry);
            }

            long offset = m_fs.Position;
            foreach (var fe in m_entries)
                fe.WriteCentralDirectoryStructureEntry(m_fs);
            long size = m_fs.Position - offset;
            m_fs.WriteUInt32(0x06054b50U);
            m_fs.WriteUInt16(0);
            m_fs.WriteUInt16(0);
            m_fs.WriteUInt16((ushort)m_entries.Count);
            m_fs.WriteUInt16((ushort)m_entries.Count);
            m_fs.WriteUInt32((uint)size);
            m_fs.WriteUInt32((uint)offset);
            m_fs.WriteUInt16(0);
        }

        private Stream m_fs;
        private string m_basePath;
        private List<FileEntry> m_entries = new List<FileEntry>();
        private HashSet<string> m_fileSet = new HashSet<string>();

        /// <summary>
        /// This class manages a file/directory.
        /// </summary>
        private sealed class FileEntry
        {
            /// <summary>
            /// Initializes an entry for a file.
            /// </summary>
            /// <param name="fileName">Path to a file/directory to the archive.</param>
            /// <param name="nameToStore">File path, which is stored in the archive.</param>
            public FileEntry(string fileName, string nameToStore)
            {
                m_fi = new FileInfo(fileName);
                m_name = nameToStore;
                m_compressedSize = IsDirectory ? 0U : (uint)m_fi.Length;
                m_data = null;
            }

            /// <summary>
            /// Initializes an entry for memory data.
            /// </summary>
            /// <param name="nameToStore">File path to store in the archive.</param>
            /// <param name="data">Content of the file.</param>
            public FileEntry(string fileName, string nameToStore, byte[] data)
            {
                m_fi = new FileInfo(fileName);
                m_name = nameToStore;
                m_data = data;
                m_compressedSize = (uint)data.Length;
            }

            /// <summary>
            /// Write local file header.
            /// </summary>
            /// <param name="s">Stream to write on.</param>
            public void WriteLocalFileHeader(Stream s)
            {
                WriteEntry(s, true);
            }

            /// <summary>
            /// Write file header on central directory structure.
            /// </summary>
            /// <param name="s">Stream to write on.</param>
            public void WriteCentralDirectoryStructureEntry(Stream s)
            {
                WriteEntry(s, false);
            }

            /// <summary>
            /// Write local file header or file header.
            /// </summary>
            /// <param name="s">Stream to write on.</param>
            /// <param name="forLocalFileHeader">To write local file header,
            /// this should be true; otherwise false.</param>
            private void WriteEntry(Stream s, bool forLocalFileHeader)
            {
                if (forLocalFileHeader)
                {
                    if (m_bodyWritten)
                        throw new ApplicationException("File is already archived.");
                }
                else
                {
                    if (!m_bodyWritten)
                        throw new ApplicationException("File is not archived!");
                }

                System.Diagnostics.Trace.WriteLine(m_name);

                bool needCompress = IsToCompress;
                if (forLocalFileHeader) m_offset = s.Position;
                s.WriteUInt32(forLocalFileHeader ? 0x04034B50U : 0x02014B50);
                if (!forLocalFileHeader)
                    s.WriteUInt16(0x2D); // made by (I don't know the actual meaning of it)
                s.WriteUInt16(IsDirectory ? 0xAU : 0x14U); // req version to extract
                s.WriteUInt16(UseUTF8 ? (1U << 11) : 0U); // Unicode flag (names are in UTF-8)
                s.WriteUInt16(needCompress ? 8U : 0U); // Deflate or not
                s.WriteUInt16(DosTime);
                s.WriteUInt16(DosDate);
                s.WriteUInt32(m_crc32); // CRC32
                s.WriteUInt32((uint)m_compressedSize);
                s.WriteUInt32(IsDirectory ? 0U : (uint)m_fi.Length);

                byte[] nameBin = StringToBytes(m_name);
                byte[] extra = new byte[0]; // empty on this code
                s.WriteUInt16((uint)nameBin.Length);
                s.WriteUInt16((uint)extra.Length);

                if (!forLocalFileHeader)
                {
                    s.WriteUInt16(0); // file comment length
                    s.WriteUInt16(0); // disk no.
                    s.WriteUInt16(0); // internal file attribute
                    s.WriteUInt32((uint)m_fi.Attributes); // external file attribute
                    s.WriteUInt32((uint)m_offset);
                }

                s.Write(nameBin);
                s.Write(extra);
                
                // Calculate LfhSize for MSIX before writing file data
                if (forLocalFileHeader)
                {
                    m_lfhSize = s.Position - m_offset;
                }

                if (!forLocalFileHeader)
                    return;

                if (IsDirectory)
                {
                    m_bodyWritten = true;
                    return;
                }

                long dataPos = s.Position;
                Stream? ws;
                Stream? ds = null;
                if (needCompress)
                    ws = ds = new DeflateStream(s, CompressionMode.Compress, true);
                else
                    ws = s;
                try
                {
                    var crc32 = new Crc32();
                    if (m_data != null)
                    {
                        ws.Write(m_data, 0, m_data.Length);
                        crc32.Append(m_data);
                    }
                    else
                    {
                        byte[] buf = new byte[1024 * 1024];
                        using (Stream fs = m_fi.OpenRead())
                        {
                            for (; ; )
                            {
                                int len = fs.Read(buf, 0, buf.Length);
                                if (len == 0)
                                    break;

                                ws.Write(buf, 0, len);
                                crc32.Append(buf.AsSpan(0, len));
                            }
                        }
                    }
                    byte[] hash = new byte[4];
                    crc32.GetCurrentHash(hash);
                    m_crc32 = BitConverter.ToUInt32(hash);
                }
                finally
                {
                    ds?.Dispose();
                }
                long lastPos = s.Position;
                m_compressedSize = lastPos - dataPos;

                // write CRC32 and the compressed size
                s.Position = m_offset + 14;
                s.WriteUInt32(m_crc32);
                s.WriteUInt32((uint)m_compressedSize);
                s.Position = lastPos;

                m_bodyWritten = true;
            }

            private FileInfo m_fi;
            private string m_name;
            private byte[]? m_data;  // For memory-based entries
            private long m_offset = 0;
            private uint m_crc32 = 0;
            private long m_compressedSize = 0;
            private long m_lfhSize = 0;
            private bool m_bodyWritten = false;
            
            /// <summary>
            /// Gets the Local File Header size for MSIX compatibility
            /// </summary>
            public long LfhSize => m_lfhSize;

            public bool IsDirectory
            {
                get
                {
                    if ((m_fi.Attributes & FileAttributes.Directory) != 0)
                        return true;
                    return false;
                }
            }

            public ushort DosDate
            {
                get
                {
                    var dt = m_fi.LastWriteTime;
                    return (ushort)(((dt.Year - 1980) << 9) | (dt.Month << 5) | dt.Day);
                }
            }

            public ushort DosTime
            {
                get
                {
                    var dt = m_fi.LastWriteTime;
                    return (ushort)((dt.Hour << 11) | (dt.Minute << 5) | (dt.Second >> 1));
                }
            }

            /// <summary>
            /// Whether or not to compress the file.
            /// </summary>
            /// <remarks>
            /// Decision is made regarding the extension, file size and other
            /// file properties.</remarks>
            private bool IsToCompress
            {
                get
                {
                    if (IsDirectory)
                        return false;

                    // The file is too small to compress; compression may
                    // make the file larger.
                    if (m_fi.Length < 256)
                        return false;

                    // Some file types are natively compressed and deflate cannot
                    // compress them efficiently; no compression is a good choice
                    // for such files.
                    string ext = Path.GetExtension(m_name).ToLower();
                    switch (ext)
                    {
                        case ".jpg":
                        case ".zip":
                        case ".docx":
                        case ".xlsx":
                        case ".pptx":
                            return false;
                        default:
                            return true;
                    }
                }
            }

            private byte[] StringToBytes(string str)
            {
                if (UseUTF8)
                    return Encoding.UTF8.GetBytes(str);
                return Encoding.Default.GetBytes(str); // use the locale setting rather than UTF-8
            }

            /// <summary>
            /// Prefers UTF-8 for file name encoding.
            /// </summary>
            public bool UseUTF8 = true;
            
            /// <summary>
            /// Gets the uncompressed size of the file
            /// </summary>
            public long UncompressedSize => IsDirectory ? 0 : m_fi.Length;
            
            /// <summary>
            /// Gets the name of the entry in the archive
            /// </summary>
            public string Name => m_name;
        }
    }

    /// <summary>
    /// BinaryWriter like extension methods for Stream class.
    /// </summary>
    public static class BinaryWriterStreamHelper
    {
        public static void Write(this Stream s, byte[] data)
        {
            s.Write(data, 0, data.Length);
        }

        public static void WriteUInt32(this Stream s, uint v)
        {
            s.Write(BitConverter.GetBytes(v));
        }

        public static void WriteUInt16(this Stream s, uint v)
        {
            s.Write(BitConverter.GetBytes((ushort)v));
        }
    }
}