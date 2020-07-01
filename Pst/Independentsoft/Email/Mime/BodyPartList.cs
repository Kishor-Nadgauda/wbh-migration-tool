using System;
using System.Collections.Generic;

namespace Independentsoft.Email.Mime
{
    /// <summary>
    /// Class BodyPartList.
    /// </summary>
    public class BodyPartList : List<BodyPart>
    {
        /// <summary>
        /// Adds the specified attachment.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        public void Add(Attachment attachment)
        {
            base.Add(new BodyPart(attachment));
        }
    }
}
