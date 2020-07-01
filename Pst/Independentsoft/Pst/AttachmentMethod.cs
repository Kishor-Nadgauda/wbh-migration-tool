using System;

namespace Independentsoft.Pst
{
    public enum AttachmentMethod
    {
        NoAttachment,
        AttachByValue,
        AttachByReference,
        AttachByReferenceResolve,
        AttachByReferenceOnly,
        EmbeddedMessage,
        Ole,
        None
    }
}
