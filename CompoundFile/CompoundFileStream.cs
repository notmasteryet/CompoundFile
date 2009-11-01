// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace CompoundFile
{
    /// <summary>
    /// Base stream for retrival data from a Compound File.
    /// </summary>
    abstract class CompoundFileStream : Stream
    {
        CompoundFileStorage storage;

        internal CompoundFileStorage Storage { get { return storage; } }

        internal CompoundFileSystem System { get { return Storage.System; } }

        long position;
        long length;
        int pageSize;

        internal int PageSize { get { return pageSize; } }
        
        byte[] page = null;
        int pageIndex = -1;

        internal CompoundFileStream(CompoundFileStorage storage, int pageSize)
        {
            this.storage = storage;
            this.position = 0;
            this.length = storage.Length;
            this.pageSize = pageSize;
        }

        public override bool CanRead
        {
            get { return true; }
        }

        public override bool CanSeek
        {
            get { return true; }
        }

        public override bool CanWrite
        {
            get { return false; }
        }

        public override void Flush()
        {
            throw new NotImplementedException();
        }

        public override long Length
        {
            get { return length; }
        }

        public override long Position
        {
            get { return position; } 
            set { Seek(value, SeekOrigin.Begin); }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (count <= 0) return count;

            if (position >= length) return 0;

            int canRead = count > (length - position) ? (int)(length - position) : count;

            long pageStartPosition = pageSize * pageIndex;
            long pageEndPosition = pageStartPosition;
            if (page != null) pageEndPosition += page.Length;

            if (position < pageStartPosition || position + canRead > pageEndPosition)
            {
                // whole data not in the current page
                byte[] cachedPage = this.page;
                int cachedPageIndex = this.pageIndex;

                // try all pages
                int startPageIndex = (int)(position / pageSize);
                int lastPageIndex = (int)((position + canRead - 1) / pageSize);

                if (startPageIndex != cachedPageIndex)
                {
                    this.page = GetPageData(startPageIndex);
                    this.pageIndex = startPageIndex;
                }
                int startPageOffset = (int)(position - startPageIndex * pageSize);
                int startPageTail = page.Length - startPageOffset;
                if (startPageTail >= canRead)
                {
                    Array.Copy(page, startPageOffset, buffer, offset, canRead);
                    position += canRead;
                }
                else
                {
                    Array.Copy(page, startPageOffset, buffer, offset, startPageTail);
                    offset += startPageTail;
                    position += startPageTail;
                    int leftToRead = canRead - startPageTail;
                    for (int i = startPageIndex + 1; i < lastPageIndex; i++)
                    {
                        if (cachedPageIndex == i)
                        {
                            this.page = cachedPage;
                            this.pageIndex = cachedPageIndex;
                        }
                        else
                        {
                            this.page = GetPageData(i);
                            this.pageIndex = i;
                        }
                        Array.Copy(page, 0, buffer, offset, page.Length);
                        offset += page.Length;
                        position += page.Length;
                        leftToRead -= page.Length;
                    }

                    if (cachedPageIndex == lastPageIndex)
                    {
                        this.page = cachedPage;
                        this.pageIndex = cachedPageIndex;
                    }
                    else
                    {
                        this.page = GetPageData(lastPageIndex);
                        this.pageIndex = lastPageIndex;
                    }

                    Array.Copy(page, 0, buffer, offset, leftToRead);
                    position += leftToRead;
                }
            }
            else // whole data in the page
            {
                Array.Copy(page, (int)(position - pageStartPosition), buffer, offset, canRead);
                position += canRead;
            }

            return canRead;
        }

        protected abstract byte[] GetPageData(int pageIndex);

        public override long Seek(long offset, SeekOrigin origin)
        {
            switch (origin)
            {
                case SeekOrigin.Begin:
                    position = offset;
                    break;
                case SeekOrigin.Current:
                    position += offset;
                    break;
                case SeekOrigin.End:
                    position = Length + offset;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("origin");
            }

            return position;
        }

        public override void SetLength(long value)
        {
            throw new NotImplementedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotImplementedException();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                storage = null;    // lost the storage and system
                page = null;
            }
        }
    }

    class FATStream : CompoundFileStream 
    {
        internal FATStream(CompoundFileStorage storage)
            : base(storage, storage.System.GetSectorSize())
        {
        }

        protected override byte[] GetPageData(int pageIndex)
        {
            byte[] page = new byte[PageSize];
            uint sector = System.GetStreamNextSector(
                Storage.Entry.StartingSectorLocation, pageIndex);
            return ReaderUtils.ReadFragment(System.BaseStream,
                System.GetSectorOffset(sector), PageSize);
        }
    }

    class MiniFATStream : CompoundFileStream
    {
        internal MiniFATStream(CompoundFileStorage storage)
            : base(storage, storage.System.GetMiniSectorSize())
        {
        }

        protected override byte[] GetPageData(int pageIndex)
        {
            byte[] page = new byte[PageSize];
            uint sector = System.GetMiniStreamNextSector(
                Storage.Entry.StartingSectorLocation, pageIndex);
            return ReaderUtils.ReadFragment(System.BaseStream,
                System.GetMiniSectorOffset(sector), PageSize);
        }
    }
}
