using System;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class ParameterList.
    /// </summary>
    public class ParameterList : List<Parameter>
    {
        /// <summary>
        /// Removes the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        public void Remove(string name)
        {
            if (name != null)
            {
                for (int i = this.Count - 1; i >= 0; i--)
                {
                    if (this[i] != null && this[i].Name != null && this[i].Name.ToLower() == name.ToLower())
                    {
                        this.RemoveAt(i);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the <see cref="Parameter"/> with the specified name.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>Parameter.</returns>
        public Parameter this[string name]
        {
            get
            {
                if (name != null)
                {
                    for (int i = 0; i < this.Count; i++)
                    {
                        if (this[i] != null && this[i].Name != null && this[i].Name.ToLower() == name.ToLower())
                        {
                            return this[i];
                        }
                    }
                }

                return null;
            }
        }
    }
}
