using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ClosedXML.Report.Options
{
    public class TagsRegister
    {
        private static readonly Dictionary<string, TagRegItem> PluginTypes = new Dictionary<string, TagRegItem>();

        public static OptionTag CreateOption(string name, Dictionary<string, string> parameters)
        {
            name = name.ToLower();
            TagRegItem typeReg;
            lock (PluginTypes)
            {
                if (!PluginTypes.ContainsKey(name))
                {
                    Debug.WriteLine(string.Format("Unknown option tag {0}.", name));
                    return null;
                }
                typeReg = PluginTypes[name];
            }
            var optionTag = (OptionTag)Activator.CreateInstance(typeReg.TagType);
            optionTag.Parameters = parameters;
            optionTag.Name = name;
            return optionTag;
        }

        public static void Add(Type tagType, string name)
        {
            if (!typeof(OptionTag).IsAssignableFrom(tagType))
                throw new ArgumentException("Type of tagType should be assignable from IOptionTag.");

            lock (PluginTypes)
            {
                PluginTypes[name.ToLower()] = new TagRegItem(tagType);
            }
        }

        public static void Add<T>(string name) where T : OptionTag
        {
            Add(typeof(T), name);
        }

        protected class TagRegItem
        {
            public TagRegItem(Type tagType)
            {
                TagType = tagType;
            }

            public Type TagType { get; private set; }
        }
    }
}