// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CompoundFile
{
    /// <summary>
    /// File System that base on Compound File structure. The class 
    /// provides methods read the structure and the data of the
    /// Compound File.
    /// 
    /// See Microsoft's [MS-CFB] "Compound File Binary File Format" 
    /// reference for more details.
    /// </summary>
    public class CompoundFileSystem : IDisposable
    {
        string filename;
        Stream baseStream;        
        bool disposed;

        CompoundFileStorage rootStorage = null;

        CFHeader header;

        /// <summary>
        /// Gets base stream.
        /// </summary>
        public Stream BaseStream
        {
            get { return baseStream; }
        }

        internal CFHeader Header
        {
            get { return header; }
        }

        internal bool IsVersion3
        {
            get { return Header.MajorVersion == 3; }
        }

        /// <summary>
        /// Gets Compound File creation date.
        /// </summary>
        public DateTime Created
        {
            get
            {
                return String.IsNullOrEmpty(filename) ?
                    default(DateTime) : File.GetCreationTimeUtc(filename);
            }
        }

        /// <summary>
        /// Gets Compound File modification date.
        /// </summary>
        public DateTime Modified
        {
            get
            {
                return String.IsNullOrEmpty(filename) ?
                    default(DateTime) : File.GetLastWriteTimeUtc(filename);
            }
        }

        /// <summary>
        /// Creates Compound File system object.
        /// </summary>
        /// <param name="filename">File path.</param>
        public CompoundFileSystem(string filename)
        {
            if (filename == null) throw new ArgumentNullException("filename");

            this.filename = filename;
            this.baseStream = File.OpenRead(filename);
            Initialize();
        }

        /// <summary>
        /// Creates Compound File system object.
        /// </summary>
        /// <param name="stream">Stream w/random access.</param>
        public CompoundFileSystem(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");

            this.baseStream = stream;
            Initialize();
        }

        /// <summary>
        /// Deinitializes the object.
        /// </summary>
        ~CompoundFileSystem()
        {
            Dispose(false);
        }

        private void Initialize()
        {
            header = ReaderUtils.ReadHeader(BaseStream);
            ReaderUtils.ValidateHeader(header);
        }

        /// <summary>
        /// Disposes the object.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposes the object.
        /// </summary>
        /// <param name="disposing">True if call via public Dispose method.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) return;

            disposed = true;
            baseStream.Close();
        }

        /// <summary>
        /// Gets Root storage.
        /// </summary>
        /// <returns>Root storage.</returns>
        public CompoundFileStorage GetRootStorage()
        {
            if (rootStorage == null)
            {
                const int RootDirectoryStreamID = 0;
                rootStorage = new CompoundFileStorage(this, RootDirectoryStreamID, new uint[0]);

                ValidateRootStorage();
            }
            return rootStorage;
        }

        private void ValidateRootStorage()
        {
            if (rootStorage.ObjectType != CompoundFileObjectType.Root)
            {
                rootStorage = null;
                throw new CompoundFileException("Invalid root storage");
            }
        }

        internal long GetSectorOffset(uint sectorNumber)
        {
            return ((long)sectorNumber + 1) << header.SectorShift;
        }

        internal int GetSectorSize()
        {
            return 1 << header.SectorShift;
        }

        internal int GetMiniSectorSize()
        {
            return 1 << header.MiniSectorShift;
        }

        internal long GetMiniSectorOffset(uint sectorNumber)
        {
            return ToPhysicalStreamOffset(GetRootStorage().Entry.StartingSectorLocation,
                sectorNumber << header.MiniSectorShift);
        }

        internal long ToPhysicalStreamOffset(uint firstSector, long offsetWithinLogicalStream)
        {
            int sectorIndex = checked((int)(offsetWithinLogicalStream >> header.SectorShift));
            if (sectorIndex == 0)
            {
                return GetSectorOffset(firstSector) + offsetWithinLogicalStream;
            }

            uint sector = GetStreamNextSector(firstSector, sectorIndex);
            int mask = (1 << Header.SectorShift) - 1;
            return GetSectorOffset(sector) + (offsetWithinLogicalStream & mask);
        }

        internal uint GetStreamNextSector(uint sector, int iterations)
        {
            for (int i = 0; i < iterations; i++)
            {
                // sector * 4 / sectorSize
                int fatSectorIndex = (int)((sector >> (Header.SectorShift - 2)));
                
                int mask = (1 << (Header.SectorShift - 2)) - 1;
                int entryIndex = (int)(sector & mask);

                uint fatSector = GetFATSector(fatSectorIndex);

                sector = ReaderUtils.ReadUInt32(BaseStream,
                    GetSectorOffset(fatSector) + (entryIndex << 2));

                if (sector > FATSectorIds.MAXREGSECT)
                    throw new CompoundFileException("Short chain");
            }
            return sector; 
        }

        internal uint GetMiniStreamNextSector(uint sector, int iterations)
        {
            for (int i = 0; i < iterations; i++)
            {
                // sector * 4 / sectorSize
                long entryOffset = (long)sector << 2;
                if (entryOffset >> Header.SectorShift >= Header.MiniFATSectorsCount)
                    throw new CompoundFileException("Mini FAT sector index out of range");

                long offset = ToPhysicalStreamOffset(Header.FirstMiniFATSectorLocation, entryOffset);
                sector = ReaderUtils.ReadUInt32(BaseStream, offset);

                if (sector > FATSectorIds.MAXREGSECT)
                    throw new CompoundFileException("Short chain");
            }
            return sector;
        }

        private uint GetFATSector(int fatSectorIndex)
        {
            if (fatSectorIndex < ReaderUtils.HeaderDIFATSectorsCount)
            {
                return Header.DIFAT[fatSectorIndex];
            }
            else
            {
                int itemsPerPage = (GetSectorSize() >> 2) - 1;
                int pageIndex = (fatSectorIndex - ReaderUtils.HeaderDIFATSectorsCount) / itemsPerPage;
                if (pageIndex >= Header.FATSectorsCount)
                    throw new CompoundFileException("FAT sector index out of range");

                uint currentSector = Header.FirstDIFATSectorLocation;
                for (int i = 0; i < pageIndex; i++)
                {
                    // last item is reference to next page
                    long offset = GetSectorOffset(currentSector) +
                        (itemsPerPage << 2);
                    currentSector = ReaderUtils.ReadUInt32(BaseStream, offset);

                }
                int entryIndex = (fatSectorIndex - ReaderUtils.HeaderDIFATSectorsCount) % itemsPerPage;
                return ReaderUtils.ReadUInt32(BaseStream,
                    GetSectorOffset(currentSector) + (entryIndex << 2));
            }
        }
    }
}
