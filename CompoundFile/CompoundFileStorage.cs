// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CompoundFile
{
    /// <summary>
    /// Compound File Storage.
    /// </summary>
    public class CompoundFileStorage
    {
        CompoundFileSystem system;
        uint streamID;
        CFDirectoryEntry entry;
        IEnumerable<uint> ancestors;

        internal CompoundFileSystem System { get { return system; } }
        internal CFDirectoryEntry Entry { get { return entry; } }

        /// <summary>
        /// Gets extended of the storage.
        /// </summary>
        public ExtendedName Name
        {
            get { return new ExtendedName(entry.Name, 0, entry.Name.Length - 1); }
        }

        /// <summary>
        /// Gets storage CLSID.
        /// </summary>
        public Guid Guid
        {
            get { return entry.CLSID; }
        }

        /// <summary>
        /// Gets storage stream length.
        /// </summary>
        public long Length
        {
            get { return (long)entry.StreamSize; }
        }

        /// <summary>
        /// Gets storage creation date.
        /// </summary>
        public DateTime Created
        {
            get { return entry.CreationTime == 0 ?
                System.Created : FromFILETIME(entry.CreationTime); }
        }

        /// <summary>
        /// Gets storage modification date.
        /// </summary>
        public DateTime Modified
        {
            get { return entry.CreationTime == 0 ?
                System.Modified : FromFILETIME(entry.ModifiedTime); }
        }

        /// <summary>
        /// Gets storage object type.
        /// </summary>
        public CompoundFileObjectType ObjectType
        {
            get { return (CompoundFileObjectType)entry.ObjectType; }
        }

        internal CompoundFileStorage(CompoundFileSystem system, uint streamID, IEnumerable<uint> ancestors)
        {
            this.system = system;
            this.streamID = streamID;
            this.ancestors = ancestors;

            Initialize();
        }

        private void Initialize()
        {
            entry = ReaderUtils.ReadDirectoryEntry(System.BaseStream,
                GetStreamOffset(), System.IsVersion3);
            ReaderUtils.ValidateDirectoryEntry(entry);
        }

        private long GetStreamOffset()
        {
            long offsetWithinLogicalStream = streamID * ReaderUtils.DirectoryEntrySize;

            if (!System.IsVersion3 &&
                offsetWithinLogicalStream >> System.Header.SectorShift >= System.Header.DirectorySectorsCount)
            {

                throw new CompoundFileException("Stream ID out of range");
            }

            return System.ToPhysicalStreamOffset(System.Header.FirstDirectorySectorLocation, 
                offsetWithinLogicalStream);
        }

        private void AppendChild(List<CompoundFileStorage> collection, uint streamID, Dictionary<uint, CompoundFileStorage> processed, IEnumerable<uint> ancestors)
        {
            if (processed.ContainsKey(streamID))
                throw new CompoundFileException("Circular structure: siblings");

            CompoundFileStorage item = new CompoundFileStorage(System, streamID, ancestors);
            processed.Add(streamID, item);

            if (item.Entry.LeftSiblingID != DirectoryStreamIds.NOSTREAM)
            {
                AppendChild(collection, item.Entry.LeftSiblingID, processed, ancestors);
            }
            if (item.ObjectType != CompoundFileObjectType.Unknown)
            {
                // skiping unknown
                collection.Add(item);
            }
            if (item.Entry.RightSiblingID != DirectoryStreamIds.NOSTREAM)
            {
                AppendChild(collection, item.Entry.RightSiblingID, processed, ancestors);
            }
        }

        /// <summary>
        /// Gets all children storages.
        /// </summary>
        /// <returns>Children storages.</returns>
        public IEnumerable<CompoundFileStorage> GetStorages()
        {
            List<CompoundFileStorage> entries = new List<CompoundFileStorage>();
            Dictionary<uint, CompoundFileStorage> processed = new Dictionary<uint, CompoundFileStorage>();
            uint childID = entry.ChildID;
            if (childID != DirectoryStreamIds.NOSTREAM)
            {
                List<uint> newAncestors = new List<uint>();
                newAncestors.Add(streamID);
                newAncestors.AddRange(ancestors);

                if(newAncestors.Contains(childID))
                    throw new CompoundFileException("Circular structure: ancestors");

                AppendChild(entries, childID, processed, ancestors);
            }
            return entries;
        }

        /// <summary>
        /// Creates the stream to access/read the data.
        /// </summary>
        /// <returns>Stream object.</returns>
        public Stream CreateStream()
        {
            if (entry.ObjectType == DirectoryObjectTypes.Stream)
            {
                if (entry.StreamSize > (ulong)System.BaseStream.Length)
                    throw new CompoundFileException("Stream length");

                if (entry.StreamSize < System.Header.MiniStreamCutoffSize)
                    return new MiniFATStream(this);
                else
                    return new FATStream(this);
            }
            else
                throw new NotSupportedException();
        }

        private DateTime FromFILETIME(ulong time)
        {
            return ReaderUtils.FromFiletime(time);
        }

        /// <summary>
        /// Copies whole content of the stream into local file.
        /// Creation and modification date will be set as well.
        /// </summary>
        /// <param name="path">Destination file path.</param>
        public void CopyToFile(string path)
        {
            using (Stream source = CreateStream(),
                target = File.Create(path))
            {
                int read;
                const int BufferSize = 8192;
                byte[] buffer = new byte[BufferSize];
                while ((read = source.Read(buffer, 0, BufferSize)) > 0)
                {
                    target.Write(buffer, 0, read);
                }
            }
            File.SetCreationTimeUtc(path, Created);
            File.SetLastWriteTimeUtc(path, Modified);
        }
    }
}
