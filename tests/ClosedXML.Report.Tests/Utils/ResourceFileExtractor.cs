using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ClosedXML.Report.Tests.Utils
{
    /// <summary>
    /// Summary description for ResourceFileExtractor.
    /// </summary>
    public sealed class ResourceFileExtractor
    {
        #region Static
        #region Private fields
        private static readonly Dictionary<string, ResourceFileExtractor> DefaultExtractors = new();
        #endregion
        #region Public properties
        /// <summary>Instance of resource extractor for executing assembly </summary>
        public static ResourceFileExtractor Instance
        {
            get
            {
                Assembly assembly = Assembly.GetCallingAssembly();
                string key = assembly.GetName().FullName;
                if (!DefaultExtractors.TryGetValue(key, out var resourceFileExtractor))
                {
                    lock (DefaultExtractors)
                    {
                        if (!DefaultExtractors.TryGetValue(key, out resourceFileExtractor))
                        {
                            resourceFileExtractor = new ResourceFileExtractor(assembly, true, null);
                            DefaultExtractors.Add(key, resourceFileExtractor);
                        }
                    }
                }
                return resourceFileExtractor;
            }
        }
        #endregion
        #endregion
        #region Private fields
        private readonly Assembly _assembly;
        private readonly ResourceFileExtractor _baseExtractor;
        private readonly string _assemblyName;

        private bool _isStatic;
        private string _resourceFilePath;
        #endregion
        #region Constructors
        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="resourceFilePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(string resourceFilePath, ResourceFileExtractor baseExtractor)
                : this(Assembly.GetCallingAssembly(), baseExtractor)
        {
            _resourceFilePath = resourceFilePath;
        }
        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(ResourceFileExtractor baseExtractor)
                : this(Assembly.GetCallingAssembly(), baseExtractor)
        {
        }
        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="resourcePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
        public ResourceFileExtractor(string resourcePath)
                : this(Assembly.GetCallingAssembly(), resourcePath)
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="resourcePath"></param>
        public ResourceFileExtractor(Assembly assembly, string resourcePath)
                : this(assembly ?? Assembly.GetCallingAssembly())
        {
            _resourceFilePath = resourcePath;
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        public ResourceFileExtractor()
                : this(Assembly.GetCallingAssembly())
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        public ResourceFileExtractor(Assembly assembly)
                : this(assembly ?? Assembly.GetCallingAssembly(), (ResourceFileExtractor) null)
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(Assembly assembly, ResourceFileExtractor baseExtractor)
                : this(assembly ?? Assembly.GetCallingAssembly(), false, baseExtractor)
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="isStatic"></param>
        /// <param name="baseExtractor"></param>
        /// <exception cref="ArgumentNullException">Argument is null.</exception>
        private ResourceFileExtractor(Assembly assembly, bool isStatic, ResourceFileExtractor baseExtractor)
        {
            #region Check
            if (ReferenceEquals(assembly, null))
            {
                throw new ArgumentNullException("assembly");
            }
            #endregion
            _assembly = assembly;
            _baseExtractor = baseExtractor;
            _assemblyName = Assembly.GetName().Name;
            IsStatic = isStatic;
            _resourceFilePath = ".Resources.";
        }
        #endregion
        #region Public properties
        /// <summary> Work assembly </summary>
        public Assembly Assembly
        {
            [DebuggerStepThrough]
            get { return _assembly; }
        }
        /// <summary> Work assembly name </summary>
        public string AssemblyName
        {
            [DebuggerStepThrough]
            get { return _assemblyName; }
        }
        /// <summary>
        /// Path to read resource files. Example: .Resources.Upgrades.
        /// </summary>
        public string ResourceFilePath
        {
            [DebuggerStepThrough]
            get { return _resourceFilePath; }
            [DebuggerStepThrough]
            set { _resourceFilePath = value; }
        }
        public bool IsStatic
        {
            [DebuggerStepThrough]
            get { return _isStatic; }
            [DebuggerStepThrough]
            set { _isStatic = value; }
        }
        public IEnumerable<string> GetFileNames()
        {
            string path = AssemblyName + _resourceFilePath;
            foreach (string resourceName in Assembly.GetManifestResourceNames())
            {
                if (resourceName.StartsWith(path))
                {
                    yield return resourceName.Replace(path, string.Empty);
                }
            }
        }
        #endregion
        #region Public methods
        public string ReadFileFromRes(string fileName)
        {
            Stream stream = ReadFileFromResToStream(fileName);
            string result;
            StreamReader sr = new StreamReader(stream);
            try
            {
                result = sr.ReadToEnd();
            }
            finally
            {
                sr.Close();
            }
            return result;
        }

        public string ReadFileFromResFormat(string fileName, params object[] formatArgs)
        {
            return string.Format(ReadFileFromRes(fileName), formatArgs);
        }

        /// <summary>
        /// Read file in current assembly by specific path
        /// </summary>
        /// <param name="specificPath">Specific path</param>
        /// <param name="fileName">Read file name</param>
        /// <returns></returns>
        public string ReadSpecificFileFromRes(string specificPath, string fileName)
        {
            ResourceFileExtractor ext = new ResourceFileExtractor(Assembly, specificPath);
            return ext.ReadFileFromRes(fileName);
        }
        /// <summary>
        /// Read file in current assembly by specific file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"><c>ApplicationException</c>.</exception>
        public Stream ReadFileFromResToStream(string fileName)
        {
            string nameResFile = AssemblyName + _resourceFilePath + fileName;
            Stream stream = Assembly.GetManifestResourceStream(nameResFile);
            #region Not found
            if (ReferenceEquals(stream, null))
            {
                #region Get from base extractor
                if (!ReferenceEquals(_baseExtractor, null))
                {
                    return _baseExtractor.ReadFileFromResToStream(fileName);
                }
                #endregion
                throw new ApplicationException("Can't find resource file " + nameResFile);
            }
            #endregion
            return stream;
        }
        #endregion
    }
}