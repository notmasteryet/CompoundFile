// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;
using CompoundFile;
using System.IO;

namespace CompoundFileUtils
{
    /// <summary>
    /// Example of using CompoundFile library.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return;
            }

            try
            {
                switch (args[0].ToLowerInvariant())
                {
                    case "-list":
                        ListContent(args[1]);
                        break;
                    case "-extract":
                        ExtractContext(args[1], args[2], args[3]);
                        break;
                    default:
                        throw new InvalidOperationException("Unknown option: " + args[0]);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }
        }

        private static void ExtractContext(string filename, string itemName, string target)
        {
            Console.WriteLine("Extracting stream \'{1}\' from of '{0}'", filename, itemName);
            CompoundFileSystem compoundFile = new CompoundFileSystem(filename);
            CompoundFileStorage storage = null;
            foreach (KeyValuePair<string, CompoundFileStorage> pair in
                ListStorage(compoundFile.GetRootStorage(), "/"))
            {
                if (pair.Key == itemName) storage = pair.Value;
            }
            if (storage != null)
            {
                storage.CopyToFile(target);
                Console.WriteLine("The {0} bytes were saved to '{1}'", storage.Length, target);
            }
            else
                Console.WriteLine("The stream was not found");
        }

        private static void ListContent(string filename)
        {
            CompoundFileSystem compoundFile = new CompoundFileSystem(filename);
            Console.WriteLine("Content of '{0}': ", filename);
            Console.WriteLine();
            Console.WriteLine("Type\tModified\tLength\tName");
            int i = 0;
            foreach (KeyValuePair<string, CompoundFileStorage> pair in
                ListStorage(compoundFile.GetRootStorage(), "/"))
            {
                Console.WriteLine("{0}\t{1}\t{2}\t\"{3}\"",
                    pair.Value.ObjectType, pair.Value.Modified, pair.Value.Length, pair.Key);
                ++i;
            }
            Console.WriteLine("{0} item(s) found", i);
        }

        private static void PrintUsage()
        {
            Console.WriteLine("USAGE: CompoundFileUtils.exe -list <compound-file>");
            Console.WriteLine("       CompoundFileUtils.exe -extract <compound-file> <stream-name> <output-file>");
        }

        /// <summary>
        /// Returns pair of item name and storage.
        /// </summary>
        /// <param name="storage">A Compound File storage.</param>
        /// <param name="parentPath">Name prefix.</param>
        /// <returns>Pairs of item name and storage.</returns>
        private static IEnumerable<KeyValuePair<string, CompoundFileStorage>> ListStorage(CompoundFileStorage storage, string parentPath)
        {
            string name = storage.Name.ToEscapedString();
            yield return new KeyValuePair<string, CompoundFileStorage>(parentPath + name, storage);

            foreach (CompoundFileStorage child in storage.GetStorages())
            {
                foreach (KeyValuePair<string, CompoundFileStorage> result in ListStorage(child, parentPath + name + "/"))
                {
                    yield return result;
                }
            }
        }

    }
}
