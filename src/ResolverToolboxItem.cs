using System;
using System.Drawing;
using System.Drawing.Design;
using System.Linq;
using System.Runtime.Serialization;

namespace SergejDerjabkin.VSAssemblyResolver
{
    [Serializable]
    public class ResolverToolboxItem : ToolboxItem
    {
        public ResolverToolboxItem(Type toolType)
            : base(toolType)
        {
            this.Bitmap = GetImage(toolType);
        }

        private static Bitmap GetImage(Type toolType)
        {
            var tb = (ToolboxBitmapAttribute)toolType.GetCustomAttributes(typeof(ToolboxBitmapAttribute), false).FirstOrDefault();
            if (tb != null)
            {
                return (Bitmap)tb.GetImage(toolType);

            }

            return null;
        }

        protected ResolverToolboxItem(SerializationInfo info, StreamingContext context)
        {
            Deserialize(info, context);
        }
    }
}
