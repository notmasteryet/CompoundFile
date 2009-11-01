// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;

namespace CompoundFile
{
    /// <summary>
    /// Compound File storage object type.
    /// </summary>
    public enum CompoundFileObjectType
    {
        /// <summary>
        /// Unknown or unassigned.
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// Storage.
        /// </summary>
        Storage = 1,

        /// <summary>
        /// Stream.
        /// </summary>
        Stream = 2,

        /// <summary>
        /// Root Storage.
        /// </summary>
        Root = 5
    }
}
