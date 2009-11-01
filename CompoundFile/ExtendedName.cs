// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;

namespace CompoundFile
{
    /// <summary>
    /// Provides name class for Compound File Storage.
    /// </summary>
    public class ExtendedName
    {
        char[] name;

        /// <summary>
        /// Gets name as array of Unicode characters.
        /// </summary>
        public char[] Name { get { return name; } }

        internal ExtendedName(char[] name)
        {
            if (name == null) throw new ArgumentNullException("name");
            this.name = name;
        }

        /// <summary>
        /// Creates extended name object.
        /// </summary>
        /// <param name="name">Array of characters.</param>
        /// <param name="offset">Offset in the array.</param>
        /// <param name="count">Count of characters.</param>
        public ExtendedName(char[] name, int offset, int count)
        {
            if (name == null) throw new ArgumentNullException("name");
            this.name = new char[count];
            Array.Copy(name, offset, this.name, 0, count);
        }

        /// <summary>
        /// Creates extended name object.
        /// </summary>
        /// <param name="name">Regular string.</param>
        public ExtendedName(string name)
        {
            if (name == null) throw new ArgumentNullException("name");
            this.name = name.ToCharArray();
        }

        /// <summary>
        /// Gets string representation.
        /// </summary>
        /// <returns>String representation.</returns>
        public override string ToString()
        {
            return new String(name);
        }

        /// <summary>
        /// Gets character array.
        /// </summary>
        /// <returns>Character array.</returns>
        public char[] ToCharArray()
        {
            return (char[])name.Clone();
        }

        /// <summary>
        /// Returns escaped string representation. Converts all
        /// control symbols and escape symbols into their code.
        /// </summary>
        /// <remarks>
        /// Escape symbols is '\'. Code if heximal string escape,
        /// e.g. tab will be presented as '\u0009'.
        /// </remarks>
        /// <returns>Escaped string representation.</returns>
        public string ToEscapedString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in name)
            {
                if(ch < ' ' || ch == '\\')
                    sb.Append(@"\u").Append(((int)ch).ToString("X4"));
                else
                    sb.Append(ch);                    
            }
            return sb.ToString();
        }
    }
}
