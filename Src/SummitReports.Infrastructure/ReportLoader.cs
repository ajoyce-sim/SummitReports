using System;
using System.Collections.Generic;
using System.Reflection;
namespace SummitReports.Infrastructure
{
    public class ReportLoader
    {
        private ReportLoader()
        {
        }
        protected Assembly _ReportAssembly = null;
        private string assemblyPath = "";
        protected void LoadReportObjectAssembly()
        {
            var binPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace(@"file:\", "");
            assemblyPath = binPath + @"\SummitReports.Objects.dll";
            try
            {
                _ReportAssembly = Assembly.LoadFrom(assemblyPath);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public List<string> ReportClassList()
        {
            if (_ReportAssembly == null) LoadReportObjectAssembly();
            List<string> pluginClassesNames = new List<string>(3);
            Type[] possiblePlugins = _ReportAssembly.GetTypes();
            try
            {

                foreach (Type t in possiblePlugins)
                {
                    if (!t.IsAbstract && !t.IsInterface && (t.GetInterface("ISummitReport") != null))
                        pluginClassesNames.Add(t.FullName);
                }
            }
            catch (Exception ex)
            {
                pluginClassesNames.Add("Exception " + ex.ToString());
            }

            return pluginClassesNames;
        }

        public T CreateInstance<T>(string ReportObjectNameToCreate)
        {
            if (!ReportObjectNameToCreate.Contains("SummitReports.Objects.")) ReportObjectNameToCreate = "SummitReports.Objects." + ReportObjectNameToCreate;
            if (_ReportAssembly == null) LoadReportObjectAssembly();
            T obj = (T)_ReportAssembly.CreateInstance(ReportObjectNameToCreate);
            if (obj != null)
            {
                return obj;
            }
            else
            {
                throw new Exception(string.Format("Could not dynamically load object {0} from assembly {1}.  Valid Classes:{2} ", ReportObjectNameToCreate, assemblyPath, String.Join(", ", this.ReportClassList())));
            }
        } 
        private static ReportLoader instance;

        public static ReportLoader Instance
        {
            get
            {
                if (instance == null)
                {
                    try
                    {
                        instance = new ReportLoader();
                    }
                    catch (System.Exception ex)
                    {
                        throw;
                    }
                }
                return instance;
            }
        }

    }
}

