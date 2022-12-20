// Copyright (c) Bodoconsult EDV-Dienstleistungen GmbH. All rights reserved.


using System.IO;
using System.Reflection;

namespace Bodoconsult.Core.Pdf.Test.Helpers
{
    internal class FileHelper
    {

        /// <summary>
        /// Get a text from a embedded resource file
        /// </summary>
        /// <param name="resourceName">resource name = plain file name with out extension and path</param>
        /// <returns></returns>
        public static string GetTextResource(string resourceName)
        {

            resourceName = $"Bodoconsult.Pdf.Test.Resources.{resourceName}";

            var ass = Assembly.GetExecutingAssembly();
            var str = ass.GetManifestResourceStream(resourceName);

            if (str == null) return null;

            string s;

            using (var file = new StreamReader(str))
            {
                s = file.ReadToEnd();
            }

            return s;
        }

    }
}
