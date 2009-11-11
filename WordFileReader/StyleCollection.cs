// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordFileReader
{
    public class StyleCollection
    {
        public const string PropertyNamePrefix = "sprm";

        IEnumerable<Prl[]> prls;

        public object this[string name]
        {
            get { return Get(name); }
        }

        internal StyleCollection(IEnumerable<Prl[]> prls)
        {
            this.prls = prls;
        }

        public object Get(string name)
        {
            if (name == null) throw new ArgumentNullException("name");

            ushort sprm = SinglePropertyModifiers.GetSprmByName(name);
            foreach (Prl[] prlSet in prls)
            {
                foreach (Prl prl in prlSet)
                {
                    if (prl.sprm.sprm == sprm)
                        return SinglePropertyValue.ParseValue(sprm, prl.operand);
                }
            }
            return null;
        }

        public IEnumerable<object> GetAll(string name)
        {
            if (name == null) throw new ArgumentNullException("name");

            ushort sprm = SinglePropertyModifiers.GetSprmByName(name);
            foreach (Prl[] prlSet in prls)
            {
                foreach (Prl prl in prlSet)
                {
                    if (prl.sprm.sprm == sprm)
                        yield return SinglePropertyValue.ParseValue(sprm, prl.operand);
                }
            }
        }

        public IEnumerable<string> GetNames()
        {
            HashSet<ushort> output = new HashSet<ushort>();
            foreach (Prl[] prlSet in prls)
            {
                foreach (Prl prl in prlSet)
                {
                    if (output.Contains(prl.sprm.sprm)) continue;
                    output.Add(prl.sprm.sprm);
                    yield return SinglePropertyModifiers.map[prl.sprm.sprm];
                }
            }
        }

    }
}
